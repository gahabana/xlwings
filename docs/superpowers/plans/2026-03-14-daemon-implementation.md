# xlwings Daemon Implementation Plan

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a persistent background Python daemon on macOS that eliminates 2-3s startup latency per RunPython call by keeping the interpreter warm and pre-importing heavy modules.

**Architecture:** A Unix domain socket server (`xlwings/daemon.py`) started on workbook open and stopped on workbook close. VBA's `ExecuteMac` detects `DAEMON=1` config and routes commands to the daemon via a small Python socket client instead of spawning a new process. Fallback to normal subprocess spawn if daemon is unreachable after 10s.

**Tech Stack:** Python stdlib (`socket`, `fcntl`, `importlib`, `threading`, `subprocess`), VBA (Main.bas), AppleScript (xlwings-dev.applescript)

**Spec:** `docs/superpowers/specs/2026-03-14-daemon-design.md`

---

## File Structure

| File | Action | Responsibility |
|------|--------|----------------|
| `xlwings/daemon.py` | Create | Socket server, EXEC/PING/SHUTDOWN handling, module reload, health-check thread, PID file management, CLI entry point |
| `xlwings/daemon_client.py` | Create | Thin socket client used by AppleScript to send commands and read responses |
| `xlwings/addin/Main.bas` | Modify (lines 35-133, 136-190) | Add `DAEMON` config check in `RunPython`, add `ExecuteMacDaemon` sub, add `LaunchDaemon` sub |
| `xlwings/xlwings-dev.applescript` | Modify | Add `DaemonHandler` for launching daemon and `DaemonExecHandler` for sending EXEC commands |
| `tests/test_daemon.py` | Create | Unit tests for daemon server, protocol, PID management, module reload |

---

## Chunk 1: Daemon Server Core

### Task 1: Socket server lifecycle (listen, accept, PING, SHUTDOWN)

**Files:**
- Create: `tests/test_daemon.py`
- Create: `xlwings/daemon.py`

- [ ] **Step 1: Write failing test for PING/PONG**

```python
# tests/test_daemon.py
import os
import socket
import tempfile
import threading
import time
import unittest


class TestDaemonProtocol(unittest.TestCase):
    """Test the daemon's newline-delimited text protocol over Unix domain sockets."""

    def setUp(self):
        self.socket_path = os.path.join(
            tempfile.gettempdir(),
            f"xlwings-test-{os.getpid()}.sock",
        )
        self.pid_file_path = os.path.join(
            tempfile.gettempdir(),
            f"xlwings-test-{os.getpid()}.pid",
        )
        # Clean up any leftover socket/pid files from a previous failed run
        for path in (self.socket_path, self.pid_file_path):
            if os.path.exists(path):
                os.unlink(path)

    def tearDown(self):
        for path in (self.socket_path, self.pid_file_path):
            if os.path.exists(path):
                os.unlink(path)

    def _send_command(self, command):
        """Helper: connect to daemon socket, send command, return response."""
        client_socket = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
        client_socket.settimeout(5)
        client_socket.connect(self.socket_path)
        client_socket.sendall(f"{command}\n".encode())
        response = b""
        while True:
            chunk = client_socket.recv(4096)
            if not chunk:
                break
            response += chunk
        client_socket.close()
        return response.decode().strip()

    def test_ping_returns_pong(self):
        from xlwings.daemon import DaemonServer

        server = DaemonServer(
            socket_path=self.socket_path,
            pid_file_path=self.pid_file_path,
            workbook_name="Test.xlsm",
            pythonpath="",
            app_path="/Applications/Microsoft Excel.app",
        )
        server_thread = threading.Thread(target=server.serve, daemon=True)
        server_thread.start()
        # Give server time to bind
        time.sleep(0.2)

        response = self._send_command("PING")
        self.assertEqual(response, "PONG")

        # Clean shutdown
        self._send_command("SHUTDOWN")
        server_thread.join(timeout=3)

    def test_shutdown_stops_server(self):
        from xlwings.daemon import DaemonServer

        server = DaemonServer(
            socket_path=self.socket_path,
            pid_file_path=self.pid_file_path,
            workbook_name="Test.xlsm",
            pythonpath="",
            app_path="/Applications/Microsoft Excel.app",
        )
        server_thread = threading.Thread(target=server.serve, daemon=True)
        server_thread.start()
        time.sleep(0.2)

        response = self._send_command("SHUTDOWN")
        self.assertEqual(response, "OK")
        server_thread.join(timeout=3)
        self.assertFalse(server_thread.is_alive())

    def test_unknown_command_returns_error(self):
        from xlwings.daemon import DaemonServer

        server = DaemonServer(
            socket_path=self.socket_path,
            pid_file_path=self.pid_file_path,
            workbook_name="Test.xlsm",
            pythonpath="",
            app_path="/Applications/Microsoft Excel.app",
        )
        server_thread = threading.Thread(target=server.serve, daemon=True)
        server_thread.start()
        time.sleep(0.2)

        response = self._send_command("BOGUS")
        self.assertTrue(response.startswith("ERROR:"))

        self._send_command("SHUTDOWN")
        server_thread.join(timeout=3)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_daemon.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'xlwings.daemon'`

- [ ] **Step 3: Implement DaemonServer with PING/SHUTDOWN**

```python
# xlwings/daemon.py
"""
xlwings daemon — persistent Python process for fast RunPython execution on macOS.

Listens on a Unix domain socket and executes commands sent from Excel via VBA/AppleScript.
Eliminates 2-3s interpreter startup latency by keeping Python warm between calls.
"""

import fcntl
import logging
import os
import signal
import socket
import sys
import threading

logger = logging.getLogger(__name__)

# Protocol: newline-delimited text commands over Unix domain socket
# Commands: PING, SHUTDOWN, EXEC <python_code>
# Responses: PONG, OK, ERROR: <traceback>


class DaemonServer:
    """Unix domain socket server that accepts commands from Excel/VBA."""

    def __init__(
        self,
        socket_path,
        pid_file_path,
        workbook_name,
        pythonpath,
        app_path,
    ):
        self.socket_path = socket_path
        self.pid_file_path = pid_file_path
        self.workbook_name = workbook_name
        self.pythonpath = pythonpath
        self.app_path = app_path
        self._running = False
        self._server_socket = None
        self._pid_file_fd = None

    def serve(self):
        """Main loop: bind socket, accept connections, dispatch commands."""
        self._write_pid_file()
        self._running = True

        self._server_socket = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
        # Remove stale socket file if it exists
        if os.path.exists(self.socket_path):
            os.unlink(self.socket_path)
        self._server_socket.bind(self.socket_path)
        os.chmod(self.socket_path, 0o600)
        self._server_socket.listen(1)
        self._server_socket.settimeout(1.0)  # Allow periodic check of _running flag

        logger.info("Daemon listening on %s", self.socket_path)

        while self._running:
            try:
                client_socket, _ = self._server_socket.accept()
            except socket.timeout:
                continue
            except OSError:
                break

            try:
                self._handle_client(client_socket)
            except Exception:
                logger.exception("Error handling client connection")
            finally:
                client_socket.close()

        self._cleanup()

    def _handle_client(self, client_socket):
        """Read one newline-delimited command, dispatch, send response, close."""
        client_socket.settimeout(30)
        raw_data = b""
        while b"\n" not in raw_data:
            chunk = client_socket.recv(4096)
            if not chunk:
                break
            raw_data += chunk

        command_line = raw_data.decode("utf-8").strip()
        if not command_line:
            return

        response = self._dispatch(command_line)
        client_socket.sendall(response.encode("utf-8"))
        client_socket.shutdown(socket.SHUT_WR)

    def _dispatch(self, command_line):
        """Route command to handler, return response string."""
        if command_line == "PING":
            return "PONG"
        elif command_line == "SHUTDOWN":
            self._running = False
            return "OK"
        elif command_line.startswith("EXEC "):
            return self._handle_exec(command_line[5:])
        else:
            return f"ERROR: Unknown command: {command_line}"

    def _handle_exec(self, python_code):
        """Execute Python code from EXEC command. Placeholder — Task 2 implements this."""
        return "ERROR: EXEC not yet implemented"

    def _write_pid_file(self):
        """Write PID file with fcntl.flock() for race-condition prevention."""
        self._pid_file_fd = open(self.pid_file_path, "w")
        fcntl.flock(self._pid_file_fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
        self._pid_file_fd.write(str(os.getpid()))
        self._pid_file_fd.flush()

    def _cleanup(self):
        """Remove socket and PID files on shutdown."""
        logger.info("Daemon shutting down")
        if self._server_socket:
            self._server_socket.close()
        if os.path.exists(self.socket_path):
            os.unlink(self.socket_path)
        if self._pid_file_fd:
            fcntl.flock(self._pid_file_fd, fcntl.LOCK_UN)
            self._pid_file_fd.close()
        if os.path.exists(self.pid_file_path):
            os.unlink(self.pid_file_path)

    def shutdown(self):
        """Signal the server to stop (called from health-check thread or signal handler)."""
        self._running = False
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_daemon.py -v`
Expected: 3 tests PASS

- [ ] **Step 5: Commit**

```bash
git add tests/test_daemon.py xlwings/daemon.py
git commit -m "feat(daemon): add socket server with PING/SHUTDOWN protocol"
```

---

### Task 2: EXEC command with module reload

**Files:**
- Modify: `tests/test_daemon.py`
- Modify: `xlwings/daemon.py`

- [ ] **Step 1: Write failing test for EXEC**

```python
# Add to tests/test_daemon.py TestDaemonProtocol class

    def test_exec_runs_python_code(self):
        """EXEC should execute Python code and return OK on success."""
        from xlwings.daemon import DaemonServer

        # Create a temporary user module
        module_dir = tempfile.mkdtemp()
        module_path = os.path.join(module_dir, "test_user_module.py")
        result_path = os.path.join(module_dir, "result.txt")
        with open(module_path, "w") as f:
            f.write(
                f'def main():\n    with open("{result_path}", "w") as f:\n        f.write("executed")\n'
            )

        server = DaemonServer(
            socket_path=self.socket_path,
            pid_file_path=self.pid_file_path,
            workbook_name="Test.xlsm",
            pythonpath=module_dir,
            app_path="/Applications/Microsoft Excel.app",
        )
        server_thread = threading.Thread(target=server.serve, daemon=True)
        server_thread.start()
        time.sleep(0.2)

        response = self._send_command(
            "EXEC import test_user_module; test_user_module.main()"
        )
        self.assertEqual(response, "OK")
        with open(result_path) as f:
            self.assertEqual(f.read(), "executed")

        self._send_command("SHUTDOWN")
        server_thread.join(timeout=3)

        # Cleanup temp files
        os.unlink(module_path)
        os.unlink(result_path)
        os.rmdir(module_dir)

    def test_exec_reloads_module_on_each_call(self):
        """EXEC should reload the user module so edits are picked up."""
        from xlwings.daemon import DaemonServer

        module_dir = tempfile.mkdtemp()
        module_path = os.path.join(module_dir, "reloadable.py")
        result_path = os.path.join(module_dir, "result.txt")

        # Version 1
        with open(module_path, "w") as f:
            f.write(
                f'def main():\n    with open("{result_path}", "w") as f:\n        f.write("v1")\n'
            )

        server = DaemonServer(
            socket_path=self.socket_path,
            pid_file_path=self.pid_file_path,
            workbook_name="Test.xlsm",
            pythonpath=module_dir,
            app_path="/Applications/Microsoft Excel.app",
        )
        server_thread = threading.Thread(target=server.serve, daemon=True)
        server_thread.start()
        time.sleep(0.2)

        # First call — runs v1
        response = self._send_command("EXEC import reloadable; reloadable.main()")
        self.assertEqual(response, "OK")
        with open(result_path) as f:
            self.assertEqual(f.read(), "v1")

        # Update module to v2
        with open(module_path, "w") as f:
            f.write(
                f'def main():\n    with open("{result_path}", "w") as f:\n        f.write("v2")\n'
            )

        # Second call — should pick up v2 via reload
        response = self._send_command("EXEC import reloadable; reloadable.main()")
        self.assertEqual(response, "OK")
        with open(result_path) as f:
            self.assertEqual(f.read(), "v2")

        self._send_command("SHUTDOWN")
        server_thread.join(timeout=3)

        os.unlink(module_path)
        os.unlink(result_path)
        os.rmdir(module_dir)

    def test_exec_syntax_error_returns_error_but_daemon_survives(self):
        """A syntax error in user code should return ERROR but not kill the daemon."""
        from xlwings.daemon import DaemonServer

        server = DaemonServer(
            socket_path=self.socket_path,
            pid_file_path=self.pid_file_path,
            workbook_name="Test.xlsm",
            pythonpath="",
            app_path="/Applications/Microsoft Excel.app",
        )
        server_thread = threading.Thread(target=server.serve, daemon=True)
        server_thread.start()
        time.sleep(0.2)

        # Execute bad code
        response = self._send_command("EXEC raise ValueError('test error')")
        self.assertTrue(response.startswith("ERROR:"))
        self.assertIn("test error", response)

        # Daemon should still respond to PING
        response = self._send_command("PING")
        self.assertEqual(response, "PONG")

        self._send_command("SHUTDOWN")
        server_thread.join(timeout=3)
```

- [ ] **Step 2: Run test to verify failures**

Run: `python -m pytest tests/test_daemon.py::TestDaemonProtocol::test_exec_runs_python_code tests/test_daemon.py::TestDaemonProtocol::test_exec_reloads_module_on_each_call tests/test_daemon.py::TestDaemonProtocol::test_exec_syntax_error_returns_error_but_daemon_survives -v`
Expected: FAIL — `_handle_exec` returns "ERROR: EXEC not yet implemented"

- [ ] **Step 3: Implement _handle_exec with module reload**

Replace the `_handle_exec` placeholder and add supporting code in `xlwings/daemon.py`:

```python
# Add to imports at top of xlwings/daemon.py
import importlib
import traceback

# Add to DaemonServer.__init__ (after self._server_socket = None):
        self._loaded_modules = {}  # module_name -> True (tracks which modules to reload)

# Replace _handle_exec method:
    def _handle_exec(self, python_code):
        """Execute Python code, reloading user modules first.

        The code is typically: "import foo; foo.main()"
        We parse out module names from import statements, reload them,
        then exec the full code string.
        """
        try:
            # Set sys.argv to match what xlwings scripts expect
            original_argv = sys.argv
            sys.argv = [
                "",
                f"--wb={self.workbook_name}",
                "--from_xl=1",
                f"--app={self.app_path}",
            ]

            # Ensure PYTHONPATH entries are on sys.path
            if self.pythonpath:
                for path_entry in self.pythonpath.split(";"):
                    path_entry = path_entry.strip()
                    if path_entry and path_entry not in sys.path:
                        sys.path.insert(0, path_entry)

            # Reload any previously-imported user modules before exec
            self._reload_user_modules(python_code)

            exec_globals = {"__builtins__": __builtins__}
            exec(python_code, exec_globals)
            return "OK"
        except Exception:
            return f"ERROR: {traceback.format_exc()}"
        finally:
            sys.argv = original_argv

    def _reload_user_modules(self, python_code):
        """Parse import statements from code and reload any already-loaded modules.

        This ensures the user always runs their latest saved code without
        restarting the daemon. Only reloads modules that were previously
        imported by the daemon (not stdlib/third-party).
        """
        # Extract module names from "import foo" or "import foo; foo.main()"
        for token in python_code.replace(";", "\n").split("\n"):
            token = token.strip()
            if token.startswith("import "):
                module_name = token.split()[1].split(".")[0]
                if module_name in sys.modules and module_name in self._loaded_modules:
                    try:
                        importlib.reload(sys.modules[module_name])
                        logger.debug("Reloaded module: %s", module_name)
                    except Exception:
                        logger.exception("Failed to reload module: %s", module_name)
                elif module_name in sys.modules:
                    # First time we see this module in daemon — track it for future reloads
                    # but don't reload yet (it was just freshly imported by exec)
                    pass

        # After exec, track all newly-imported user modules
        # (deferred to after exec completes — see _handle_exec)

# Add after the exec(python_code) call in _handle_exec, before return "OK":
            # Track newly-imported modules for future reloads
            for token in python_code.replace(";", "\n").split("\n"):
                token = token.strip()
                if token.startswith("import "):
                    module_name = token.split()[1].split(".")[0]
                    if module_name in sys.modules:
                        self._loaded_modules[module_name] = True
```

> **Note to implementer:** The final `_handle_exec` method should incorporate the module tracking inline. The above shows the logic split for clarity. The full method:

```python
    def _handle_exec(self, python_code):
        """Execute Python code, reloading user modules first."""
        try:
            original_argv = sys.argv
            sys.argv = [
                "",
                f"--wb={self.workbook_name}",
                "--from_xl=1",
                f"--app={self.app_path}",
            ]

            if self.pythonpath:
                for path_entry in self.pythonpath.split(";"):
                    path_entry = path_entry.strip()
                    if path_entry and path_entry not in sys.path:
                        sys.path.insert(0, path_entry)

            # Reload previously-imported user modules
            for module_name in list(self._loaded_modules):
                if module_name in sys.modules:
                    try:
                        importlib.reload(sys.modules[module_name])
                    except Exception:
                        logger.exception("Failed to reload module: %s", module_name)

            exec_globals = {"__builtins__": __builtins__}
            exec(python_code, exec_globals)

            # Track newly-imported modules for future reloads
            for token in python_code.replace(";", "\n").split("\n"):
                token = token.strip()
                if token.startswith("import "):
                    module_name = token.split()[1].split(".")[0]
                    if module_name in sys.modules:
                        self._loaded_modules[module_name] = True

            return "OK"
        except Exception:
            return f"ERROR: {traceback.format_exc()}"
        finally:
            sys.argv = original_argv
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_daemon.py -v`
Expected: 6 tests PASS

- [ ] **Step 5: Commit**

```bash
git add tests/test_daemon.py xlwings/daemon.py
git commit -m "feat(daemon): implement EXEC command with module reload"
```

---

### Task 3: PID file management and stale process cleanup

**Files:**
- Modify: `tests/test_daemon.py`
- Modify: `xlwings/daemon.py`

- [ ] **Step 1: Write failing tests for PID management**

```python
# Add to tests/test_daemon.py

class TestDaemonPidManagement(unittest.TestCase):
    """Test PID file locking and stale daemon detection."""

    def setUp(self):
        self.socket_path = os.path.join(
            tempfile.gettempdir(),
            f"xlwings-test-pid-{os.getpid()}.sock",
        )
        self.pid_file_path = os.path.join(
            tempfile.gettempdir(),
            f"xlwings-test-pid-{os.getpid()}.pid",
        )
        for path in (self.socket_path, self.pid_file_path):
            if os.path.exists(path):
                os.unlink(path)

    def tearDown(self):
        for path in (self.socket_path, self.pid_file_path):
            if os.path.exists(path):
                os.unlink(path)

    def test_pid_file_created_on_start(self):
        from xlwings.daemon import DaemonServer

        server = DaemonServer(
            socket_path=self.socket_path,
            pid_file_path=self.pid_file_path,
            workbook_name="Test.xlsm",
            pythonpath="",
            app_path="/Applications/Microsoft Excel.app",
        )
        server_thread = threading.Thread(target=server.serve, daemon=True)
        server_thread.start()
        time.sleep(0.2)

        self.assertTrue(os.path.exists(self.pid_file_path))
        with open(self.pid_file_path) as f:
            pid_content = f.read().strip()
        self.assertEqual(pid_content, str(os.getpid()))

        server.shutdown()
        server_thread.join(timeout=3)

    def test_pid_file_removed_on_shutdown(self):
        from xlwings.daemon import DaemonServer

        server = DaemonServer(
            socket_path=self.socket_path,
            pid_file_path=self.pid_file_path,
            workbook_name="Test.xlsm",
            pythonpath="",
            app_path="/Applications/Microsoft Excel.app",
        )
        server_thread = threading.Thread(target=server.serve, daemon=True)
        server_thread.start()
        time.sleep(0.2)

        server.shutdown()
        server_thread.join(timeout=3)

        self.assertFalse(os.path.exists(self.pid_file_path))
        self.assertFalse(os.path.exists(self.socket_path))

    def test_check_and_cleanup_stale_daemon(self):
        """check_stale_daemon should remove PID file for a dead process."""
        from xlwings.daemon import check_stale_daemon

        # Write a PID file with a definitely-dead PID
        with open(self.pid_file_path, "w") as f:
            f.write("999999999")

        # Should detect the stale daemon and clean up
        is_alive = check_stale_daemon(self.socket_path, self.pid_file_path)
        self.assertFalse(is_alive)

    def test_check_stale_daemon_with_alive_daemon(self):
        """check_stale_daemon should return True if daemon responds to PING."""
        from xlwings.daemon import DaemonServer, check_stale_daemon

        server = DaemonServer(
            socket_path=self.socket_path,
            pid_file_path=self.pid_file_path,
            workbook_name="Test.xlsm",
            pythonpath="",
            app_path="/Applications/Microsoft Excel.app",
        )
        server_thread = threading.Thread(target=server.serve, daemon=True)
        server_thread.start()
        time.sleep(0.2)

        is_alive = check_stale_daemon(self.socket_path, self.pid_file_path)
        self.assertTrue(is_alive)

        server.shutdown()
        server_thread.join(timeout=3)
```

- [ ] **Step 2: Run tests to verify failures**

Run: `python -m pytest tests/test_daemon.py::TestDaemonPidManagement -v`
Expected: FAIL — `ImportError: cannot import name 'check_stale_daemon'`

- [ ] **Step 3: Implement check_stale_daemon**

Add to `xlwings/daemon.py` as a module-level function:

```python
def check_stale_daemon(socket_path, pid_file_path):
    """Check if an existing daemon is alive. Returns True if alive, False if stale/dead.

    If a stale daemon is detected (PID file exists but process is dead or
    socket doesn't respond to PING), cleans up the stale PID/socket files.
    """
    if not os.path.exists(socket_path):
        # No socket = no daemon. Clean up PID file if orphaned.
        if os.path.exists(pid_file_path):
            os.unlink(pid_file_path)
        return False

    # Try PING via socket
    try:
        client_socket = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
        client_socket.settimeout(2)
        client_socket.connect(socket_path)
        client_socket.sendall(b"PING\n")
        response = b""
        while True:
            chunk = client_socket.recv(4096)
            if not chunk:
                break
            response += chunk
        client_socket.close()
        if response.decode().strip() == "PONG":
            return True
    except (ConnectionRefusedError, OSError, socket.timeout):
        pass

    # Daemon is dead — try to kill stale process and clean up
    if os.path.exists(pid_file_path):
        try:
            with open(pid_file_path) as f:
                stale_pid = int(f.read().strip())
            os.kill(stale_pid, signal.SIGTERM)
        except (ProcessLookupError, ValueError, PermissionError):
            pass  # Process already dead or PID file corrupt
        os.unlink(pid_file_path)

    if os.path.exists(socket_path):
        os.unlink(socket_path)

    # Clean up stale status file too
    status_file_path = socket_path.replace(".sock", ".status")
    if os.path.exists(status_file_path):
        os.unlink(status_file_path)

    return False
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_daemon.py -v`
Expected: All tests PASS

- [ ] **Step 5: Commit**

```bash
git add tests/test_daemon.py xlwings/daemon.py
git commit -m "feat(daemon): add PID file management and stale daemon cleanup"
```

---

### Task 4: Health-check thread

**Files:**
- Modify: `tests/test_daemon.py`
- Modify: `xlwings/daemon.py`

- [ ] **Step 1: Write failing test for health-check**

```python
# Add to tests/test_daemon.py
from unittest.mock import patch

class TestHealthCheck(unittest.TestCase):
    """Test the background health-check thread that monitors Excel."""

    def setUp(self):
        self.socket_path = os.path.join(
            tempfile.gettempdir(),
            f"xlwings-test-hc-{os.getpid()}.sock",
        )
        self.pid_file_path = os.path.join(
            tempfile.gettempdir(),
            f"xlwings-test-hc-{os.getpid()}.pid",
        )
        for path in (self.socket_path, self.pid_file_path):
            if os.path.exists(path):
                os.unlink(path)

    def tearDown(self):
        for path in (self.socket_path, self.pid_file_path):
            if os.path.exists(path):
                os.unlink(path)

    def test_daemon_shuts_down_when_excel_gone(self):
        """Health-check should trigger shutdown when Excel process disappears."""
        from xlwings.daemon import DaemonServer

        server = DaemonServer(
            socket_path=self.socket_path,
            pid_file_path=self.pid_file_path,
            workbook_name="Test.xlsm",
            pythonpath="",
            app_path="/Applications/Microsoft Excel.app",
            health_check_interval=0.3,
        )

        # Mock: Excel is NOT running
        with patch("xlwings.daemon.is_excel_running", return_value=False), \
             patch("xlwings.daemon.is_workbook_open", return_value=True):
            server_thread = threading.Thread(target=server.serve, daemon=True)
            server_thread.start()
            # Health check runs at 0.3s interval, should trigger shutdown quickly
            server_thread.join(timeout=3)
            self.assertFalse(server_thread.is_alive())

    def test_daemon_shuts_down_when_workbook_closed(self):
        """Health-check should trigger shutdown when workbook is closed."""
        from xlwings.daemon import DaemonServer

        server = DaemonServer(
            socket_path=self.socket_path,
            pid_file_path=self.pid_file_path,
            workbook_name="Test.xlsm",
            pythonpath="",
            app_path="/Applications/Microsoft Excel.app",
            health_check_interval=0.3,
        )

        with patch("xlwings.daemon.is_excel_running", return_value=True), \
             patch("xlwings.daemon.is_workbook_open", return_value=False):
            server_thread = threading.Thread(target=server.serve, daemon=True)
            server_thread.start()
            server_thread.join(timeout=3)
            self.assertFalse(server_thread.is_alive())
```

- [ ] **Step 2: Run tests to verify failures**

Run: `python -m pytest tests/test_daemon.py::TestHealthCheck -v`
Expected: FAIL — `health_check_interval` parameter not accepted / `is_excel_running` not found

- [ ] **Step 3: Implement health-check thread and monitoring functions**

Add to `xlwings/daemon.py`:

```python
# Module-level monitoring functions

def is_excel_running():
    """Check if Microsoft Excel process is running (macOS)."""
    try:
        subprocess.check_output(["pgrep", "-x", "Microsoft Excel"],
                                stderr=subprocess.DEVNULL)
        return True
    except subprocess.CalledProcessError:
        return False


def is_workbook_open(workbook_name):
    """Check if a specific workbook is open in Excel (macOS)."""
    try:
        result = subprocess.check_output(
            ["osascript", "-e",
             'tell application "Microsoft Excel" to get name of every workbook'],
            stderr=subprocess.DEVNULL,
            timeout=5,
        )
        open_workbooks = result.decode().strip().split(", ")
        return workbook_name in open_workbooks
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
        return False
```

Update `DaemonServer.__init__` to accept `health_check_interval`:

```python
    def __init__(
        self,
        socket_path,
        pid_file_path,
        workbook_name,
        pythonpath,
        app_path,
        health_check_interval=2.5,
    ):
        # ... existing assignments ...
        self.health_check_interval = health_check_interval
```

Add health-check thread start in `serve()`, right after `logger.info(...)`:

```python
        # Start health-check thread
        health_thread = threading.Thread(
            target=self._health_check_loop, daemon=True
        )
        health_thread.start()
```

Add the health-check method:

```python
    def _health_check_loop(self):
        """Periodically check that Excel is running and workbook is open."""
        while self._running:
            if not is_excel_running():
                logger.info("Excel process not found — shutting down daemon")
                self.shutdown()
                return
            if not is_workbook_open(self.workbook_name):
                logger.info("Workbook %s no longer open — shutting down daemon",
                            self.workbook_name)
                self.shutdown()
                return
            time.sleep(self.health_check_interval)
```

Add `import subprocess, time` to the imports if not already present.

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_daemon.py -v`
Expected: All tests PASS

- [ ] **Step 5: Commit**

```bash
git add tests/test_daemon.py xlwings/daemon.py
git commit -m "feat(daemon): add health-check thread monitoring Excel and workbook"
```

---

## Chunk 2: CLI Entry Point and Client

### Task 5: CLI entry point (`python -m xlwings.daemon`)

**Files:**
- Modify: `tests/test_daemon.py`
- Modify: `xlwings/daemon.py`

- [ ] **Step 1: Write failing test for CLI argument parsing**

```python
# Add to tests/test_daemon.py

class TestDaemonCLI(unittest.TestCase):
    """Test the daemon's command-line interface."""

    def test_parse_start_args(self):
        from xlwings.daemon import parse_args

        args = parse_args([
            "start",
            "--workbook", "PWandRC5.xlsm",
            "--socket", "/tmp/xlwings-daemon-abc123.sock",
            "--pidfile", "/tmp/xlwings-daemon-abc123.pid",
            "--pythonpath", "/Users/zh/projects/pwr",
            "--app", "/Applications/Microsoft Excel.app/Contents/MacOS",
        ])
        self.assertEqual(args.command, "start")
        self.assertEqual(args.workbook, "PWandRC5.xlsm")
        self.assertEqual(args.socket, "/tmp/xlwings-daemon-abc123.sock")
        self.assertEqual(args.pidfile, "/tmp/xlwings-daemon-abc123.pid")
        self.assertEqual(args.pythonpath, "/Users/zh/projects/pwr")
        self.assertEqual(args.app, "/Applications/Microsoft Excel.app/Contents/MacOS")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_daemon.py::TestDaemonCLI -v`
Expected: FAIL — `ImportError: cannot import name 'parse_args'`

- [ ] **Step 3: Implement CLI entry point**

Add to `xlwings/daemon.py`:

```python
import argparse


def parse_args(argv=None):
    """Parse daemon command-line arguments."""
    parser = argparse.ArgumentParser(
        description="xlwings daemon — persistent Python process for fast RunPython"
    )
    subparsers = parser.add_subparsers(dest="command")

    start_parser = subparsers.add_parser("start", help="Start the daemon")
    start_parser.add_argument("--workbook", required=True,
                              help="Name of the Excel workbook")
    start_parser.add_argument("--socket", required=True,
                              help="Path to Unix domain socket")
    start_parser.add_argument("--pidfile", required=True,
                              help="Path to PID file")
    start_parser.add_argument("--pythonpath", default="",
                              help="Semicolon-delimited Python paths")
    start_parser.add_argument("--app", default="",
                              help="Excel application path")

    return parser.parse_args(argv)


def main(argv=None):
    """CLI entry point for `python -m xlwings.daemon`."""
    args = parse_args(argv)

    if args.command == "start":
        # Check for existing daemon
        if check_stale_daemon(args.socket, args.pidfile):
            logger.info("Daemon already running for %s — exiting", args.workbook)
            return

        server = DaemonServer(
            socket_path=args.socket,
            pid_file_path=args.pidfile,
            workbook_name=args.workbook,
            pythonpath=args.pythonpath,
            app_path=args.app,
        )

        # Handle SIGTERM gracefully
        signal.signal(signal.SIGTERM, lambda signum, frame: server.shutdown())

        logger.info("Starting daemon for workbook: %s", args.workbook)
        server.serve()
    else:
        parse_args(["--help"])


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    main()
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_daemon.py -v`
Expected: All tests PASS

- [ ] **Step 5: Commit**

```bash
git add xlwings/daemon.py tests/test_daemon.py
git commit -m "feat(daemon): add CLI entry point for python -m xlwings.daemon"
```

---

### Task 6: Daemon client (socket communication helper)

**Files:**
- Create: `tests/test_daemon_client.py`
- Create: `xlwings/daemon_client.py`

- [ ] **Step 1: Write failing test for client**

```python
# tests/test_daemon_client.py
import os
import socket
import tempfile
import threading
import time
import unittest


class TestDaemonClient(unittest.TestCase):
    """Test the thin socket client used by AppleScript to talk to the daemon."""

    def setUp(self):
        self.socket_path = os.path.join(
            tempfile.gettempdir(),
            f"xlwings-test-client-{os.getpid()}.sock",
        )
        self.pid_file_path = os.path.join(
            tempfile.gettempdir(),
            f"xlwings-test-client-{os.getpid()}.pid",
        )
        for path in (self.socket_path, self.pid_file_path):
            if os.path.exists(path):
                os.unlink(path)

    def tearDown(self):
        for path in (self.socket_path, self.pid_file_path):
            if os.path.exists(path):
                os.unlink(path)

    def test_client_sends_exec_and_returns_response(self):
        from xlwings.daemon import DaemonServer
        from xlwings.daemon_client import send_command

        server = DaemonServer(
            socket_path=self.socket_path,
            pid_file_path=self.pid_file_path,
            workbook_name="Test.xlsm",
            pythonpath="",
            app_path="/Applications/Microsoft Excel.app",
        )
        server_thread = threading.Thread(target=server.serve, daemon=True)
        server_thread.start()
        time.sleep(0.2)

        response = send_command(self.socket_path, "PING")
        self.assertEqual(response, "PONG")

        send_command(self.socket_path, "SHUTDOWN")
        server_thread.join(timeout=3)

    def test_client_returns_error_on_connection_refused(self):
        from xlwings.daemon_client import send_command

        # No server running — should get connection error
        response = send_command("/tmp/nonexistent-socket.sock", "PING",
                                timeout=1)
        self.assertTrue(response.startswith("ERROR:"))

    def test_client_timeout_returns_timeout_error(self):
        """Client should return a specific timeout error for fallback detection."""
        from xlwings.daemon_client import send_command

        # No server running, short timeout — simulates daemon never starting
        response = send_command("/tmp/nonexistent-socket-fallback.sock", "EXEC foo()",
                                timeout=0.5, retry_interval=0.1)
        self.assertTrue(response.startswith("ERROR: Timeout"))

    def test_client_retries_on_connection_refused(self):
        """Client should retry connecting when daemon is still booting."""
        from xlwings.daemon import DaemonServer
        from xlwings.daemon_client import send_command

        server = DaemonServer(
            socket_path=self.socket_path,
            pid_file_path=self.pid_file_path,
            workbook_name="Test.xlsm",
            pythonpath="",
            app_path="/Applications/Microsoft Excel.app",
        )

        # Start server after a delay (simulates slow boot)
        def delayed_start():
            time.sleep(0.5)
            server.serve()

        server_thread = threading.Thread(target=delayed_start, daemon=True)
        server_thread.start()

        # Client should retry and eventually connect
        response = send_command(self.socket_path, "PING",
                                timeout=5, retry_interval=0.2)
        self.assertEqual(response, "PONG")

        send_command(self.socket_path, "SHUTDOWN")
        server_thread.join(timeout=3)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_daemon_client.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'xlwings.daemon_client'`

- [ ] **Step 3: Implement daemon_client.py**

```python
# xlwings/daemon_client.py
"""
Thin socket client for communicating with the xlwings daemon.

Used in two ways:
1. From AppleScript via `python -m xlwings.daemon_client` (replaces nc -U)
2. From Python code for testing and direct daemon communication

The client connects to the daemon's Unix domain socket, sends a command,
and prints the response. It retries on connection refused (daemon still booting).
"""

import socket
import sys
import time


def send_command(socket_path, command, timeout=10, retry_interval=0.2):
    """Send a command to the daemon and return the response.

    Retries connection every `retry_interval` seconds until `timeout`.
    Returns the response string, or "ERROR: ..." on failure.
    """
    deadline = time.monotonic() + timeout

    while time.monotonic() < deadline:
        try:
            client_socket = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
            client_socket.settimeout(timeout)
            client_socket.connect(socket_path)
            client_socket.sendall(f"{command}\n".encode())

            response = b""
            while True:
                chunk = client_socket.recv(4096)
                if not chunk:
                    break
                response += chunk
            client_socket.close()
            return response.decode().strip()
        except (ConnectionRefusedError, FileNotFoundError):
            time.sleep(retry_interval)
        except (OSError, socket.timeout) as exc:
            return f"ERROR: {exc}"

    return f"ERROR: Timeout connecting to daemon at {socket_path} after {timeout}s"


def main():
    """CLI entry point: python -m xlwings.daemon_client <socket_path> [--timeout N] <command>"""
    import argparse

    parser = argparse.ArgumentParser(description="Send a command to the xlwings daemon")
    parser.add_argument("socket_path", help="Path to daemon Unix socket")
    parser.add_argument("command", help="Command to send (e.g. 'PING', 'EXEC import foo; foo.main()')")
    parser.add_argument("--timeout", type=int, default=10, help="Connection timeout in seconds")
    args = parser.parse_args()

    response = send_command(args.socket_path, args.command, timeout=args.timeout)
    print(response, end="")

    # Exit with error code if response is an error
    if response.startswith("ERROR:"):
        sys.exit(1)


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_daemon_client.py -v`
Expected: All tests PASS

- [ ] **Step 5: Commit**

```bash
git add xlwings/daemon_client.py tests/test_daemon_client.py
git commit -m "feat(daemon): add socket client for VBA/AppleScript communication"
```

---

## Chunk 3: VBA and AppleScript Integration

### Task 7: AppleScript daemon handlers

**Files:**
- Modify: `xlwings/xlwings-dev.applescript`

The AppleScript file needs two new handlers that VBA will call:
1. `DaemonLaunchHandler` — starts the daemon in background
2. `DaemonExecHandler` — sends EXEC command via daemon_client.py

- [ ] **Step 1: Add DaemonLaunchHandler to AppleScript**

Add after the existing `SplitString` handler in `xlwings/xlwings-dev.applescript`:

```applescript
on DaemonLaunchHandler(ParameterString)
	-- Parameters: PythonInterpreter|WorkbookName|SocketPath|PidFilePath|PythonPath|AppPath
	set {PythonInterpreter, WorkbookName, SocketPath, PidFilePath, PythonPathArg, AppPath} to SplitString(ParameterString, "|")
	set ShellCommand to PythonInterpreter & " -m xlwings.daemon start --workbook '" & WorkbookName & "' --socket '" & SocketPath & "' --pidfile '" & PidFilePath & "' --pythonpath '" & PythonPathArg & "' --app '" & AppPath & "' &"
	try
		do shell script "source ~/.bash_profile;" & ShellCommand
	on error errMsg number errNumber
		try
			do shell script ShellCommand
		on error errMsg number errNumber
			return errMsg
		end try
	end try
	return "OK"
end DaemonLaunchHandler

on DaemonExecHandler(ParameterString)
	-- Parameters: PythonInterpreter|SocketPath|PythonCommand|Timeout
	set {PythonInterpreter, SocketPath, PythonCommand, TimeoutStr} to SplitString(ParameterString, "|")
	set ShellCommand to PythonInterpreter & " -m xlwings.daemon_client '" & SocketPath & "' --timeout " & TimeoutStr & " 'EXEC " & PythonCommand & "'"
	try
		do shell script "source ~/.bash_profile;" & ShellCommand
		return result
	on error errMsg number errNumber
		try
			return do shell script ShellCommand
		on error errMsg number errNumber
			return "ERROR: " & errMsg
		end try
	end try
end DaemonExecHandler
```

- [ ] **Step 2: Verify AppleScript syntax**

Run: `osacompile -o /dev/null /Users/zh/gd/git/xlwings/xlwings/xlwings-dev.applescript`
Expected: No errors

- [ ] **Step 3: Commit**

```bash
git add xlwings/xlwings-dev.applescript
git commit -m "feat(daemon): add AppleScript handlers for daemon launch and exec"
```

---

### Task 8: VBA modifications — daemon-aware RunPython

**Files:**
- Modify: `xlwings/addin/Main.bas` (lines 35-133 RunPython, add new subs)

This is the core integration point. We modify `RunPython` to check for `DAEMON=1` config and route to the daemon path when enabled.

- [ ] **Step 1: Add helper functions to Main.bas**

> **Hash consistency note:** The hash is computed only in VBA and passed to the daemon/client via CLI args and AppleScript parameters. Python never independently computes the hash — it receives the socket/PID paths from VBA. This avoids hash algorithm mismatch between VBA and Python.

Add after `ExecuteMac` sub (after line 190):

```vba
Function GetDaemonSocketPath(WorkbookFullName As String) As String
    ' Generate a deterministic socket path from the workbook's full path.
    ' Uses a simple hash to avoid path-length issues with $TMPDIR.
    Dim hashValue As Long
    Dim i As Integer
    hashValue = 0
    For i = 1 To Len(WorkbookFullName)
        hashValue = ((hashValue * 31) + Asc(Mid$(WorkbookFullName, i, 1))) And &H7FFFFFFF
    Next i
    GetDaemonSocketPath = Environ("TMPDIR") & "xlwings-daemon-" & CStr(hashValue) & ".sock"
End Function

Function GetDaemonPidFilePath(WorkbookFullName As String) As String
    Dim hashValue As Long
    Dim i As Integer
    hashValue = 0
    For i = 1 To Len(WorkbookFullName)
        hashValue = ((hashValue * 31) + Asc(Mid$(WorkbookFullName, i, 1))) And &H7FFFFFFF
    Next i
    GetDaemonPidFilePath = Environ("TMPDIR") & "xlwings-daemon-" & CStr(hashValue) & ".pid"
End Function
```

- [ ] **Step 2: Add LaunchDaemon sub**

Add after the hash functions:

```vba
Sub LaunchDaemon(interpreter As String, PYTHONPATH As String)
    ' Launch the daemon process in background for the active workbook.
    ' Called from RunPython on first invocation when DAEMON=1.
    #If Mac Then
    Dim SocketPath As String, PidFilePath As String
    Dim ParameterString As String, AppPath As String

    SocketPath = GetDaemonSocketPath(ActiveWorkbook.FullName)
    PidFilePath = GetDaemonPidFilePath(ActiveWorkbook.FullName)
    AppPath = Left(Application.Path, Len(Application.Path) - 4)

    ParameterString = interpreter
    ParameterString = ParameterString + "|" + ActiveWorkbook.Name
    ParameterString = ParameterString + "|" + SocketPath
    ParameterString = ParameterString + "|" + PidFilePath
    ParameterString = ParameterString + "|" + PYTHONPATH
    ParameterString = ParameterString + "|" + AppPath

    On Error Resume Next
        AppleScriptTask "xlwings-" & XLWINGS_VERSION & ".applescript", "DaemonLaunchHandler", ParameterString
    On Error GoTo 0
    #End If
End Sub
```

- [ ] **Step 3: Add ExecuteMacDaemon sub**

```vba
Sub ExecuteMacDaemon(PythonCommand As String, interpreter As String, PYTHONPATH As String)
    ' Execute a Python command via the running daemon.
    ' Falls back to normal ExecuteMac if daemon is unreachable after timeout.
    #If Mac Then
    Dim SocketPath As String, ParameterString As String
    Dim ExecResult As String

    SocketPath = GetDaemonSocketPath(ActiveWorkbook.FullName)

    ParameterString = interpreter
    ParameterString = ParameterString + "|" + SocketPath
    ParameterString = ParameterString + "|" + PythonCommand
    ParameterString = ParameterString + "|10"  ' 10 second timeout

    On Error GoTo DaemonExecFallback
        ExecResult = AppleScriptTask("xlwings-" & XLWINGS_VERSION & ".applescript", "DaemonExecHandler", ParameterString)
    On Error GoTo 0

    ' Check if daemon_client returned a timeout error — fall back to normal exec
    If Left$(ExecResult, 14) = "ERROR: Timeout" Then
        GoTo DaemonExecFallback
    End If

    ' Check result for Python errors
    If Left$(ExecResult, 6) = "ERROR:" Then
        ' Write error to log file and show via standard ShowError
        Dim LOG_FILE As String
        LOG_FILE = Environ("HOME") + "/xlwings.log"
        Dim f As Integer
        f = FreeFile
        Open LOG_FILE For Output As #f
        Print #f, Mid$(ExecResult, 8)
        Close #f
        ShowError (LOG_FILE)
    End If
    Exit Sub

DaemonExecFallback:
    ' Daemon unreachable — fall back to normal process spawn
    On Error GoTo 0
    ExecuteMac PythonCommand, interpreter, PYTHONPATH
    #End If
End Sub
```

- [ ] **Step 4: Modify RunPython to check DAEMON config**

In `Main.bas`, add the daemon check inside `RunPython` (after the PYTHONPATH construction at line ~94, before the `#If Mac Then` dispatch at line ~112):

```vba
    ' Check for daemon mode (Mac only)
    #If Mac Then
        Dim useDaemon As String
        useDaemon = GetConfig("DAEMON", "0")

        If useDaemon = "1" Or UCase(useDaemon) = "TRUE" Then
            ' Launch daemon if not already running (idempotent — daemon checks for existing)
            LaunchDaemon interpreter, PYTHONPATH

            ' Route to daemon execution path (passes PYTHONPATH for fallback to normal ExecuteMac)
            ExecuteMacDaemon PythonCommand, interpreter, PYTHONPATH
            Exit Function
        End If
    #End If
```

This goes right before the existing `#If Mac Then` block at line 112 that calls `ExecuteMac`.

- [ ] **Step 5: Commit**

```bash
git add xlwings/addin/Main.bas
git commit -m "feat(daemon): integrate daemon launch and exec into VBA RunPython"
```

---

### Task 9: Status file for daemon boot feedback

**Files:**
- Modify: `tests/test_daemon.py`
- Modify: `xlwings/daemon.py`

- [ ] **Step 1: Write failing test for status file**

```python
# Add to tests/test_daemon.py TestDaemonProtocol class

    def test_status_file_written_on_ready(self):
        """Daemon should write a .status file with 'ready' when accepting connections."""
        from xlwings.daemon import DaemonServer

        status_path = self.socket_path.replace(".sock", ".status")

        server = DaemonServer(
            socket_path=self.socket_path,
            pid_file_path=self.pid_file_path,
            workbook_name="Test.xlsm",
            pythonpath="",
            app_path="/Applications/Microsoft Excel.app",
        )
        server_thread = threading.Thread(target=server.serve, daemon=True)
        server_thread.start()
        time.sleep(0.3)

        self.assertTrue(os.path.exists(status_path))
        with open(status_path) as f:
            self.assertEqual(f.read().strip(), "ready")

        server.shutdown()
        server_thread.join(timeout=3)

        # Status file should be cleaned up
        self.assertFalse(os.path.exists(status_path))
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_daemon.py::TestDaemonProtocol::test_status_file_written_on_ready -v`
Expected: FAIL — no status file created

- [ ] **Step 3: Implement status file**

In `xlwings/daemon.py`, add status file path to `__init__`:

```python
        self.status_file_path = self.socket_path.replace(".sock", ".status")
```

In `serve()`, write status file right after `self._server_socket.listen(1)` and before the main loop:

```python
        # Signal readiness
        with open(self.status_file_path, "w") as f:
            f.write("ready")
```

In `_cleanup()`, add status file removal:

```python
        if os.path.exists(self.status_file_path):
            os.unlink(self.status_file_path)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_daemon.py -v`
Expected: All tests PASS

- [ ] **Step 5: Commit**

```bash
git add tests/test_daemon.py xlwings/daemon.py
git commit -m "feat(daemon): write status file on ready for boot diagnostics"
```

---

## Chunk 4: Workbook Lifecycle Events

### Task 10: VBA Workbook_Open and Workbook_BeforeClose hooks

**Files:**
- Modify: `xlwings/addin/Main.bas`

The xlwings addin doesn't currently have `Workbook_Open` / `Workbook_BeforeClose` handlers. These need to be added so the daemon starts and stops automatically with the workbook.

> **Note:** VBA event handlers (`Workbook_Open`, `Workbook_BeforeClose`) must live in the `ThisWorkbook` module of the `.xlam` addin, not in `Main.bas`. However, since xlwings uses `Main.bas` as the primary module and the addin's `ThisWorkbook` code isn't in this repo's VBA exports, we'll add callable subs to `Main.bas` that the user can call from their workbook's `ThisWorkbook` module, or that can be wired up via the addin's `ThisWorkbook`.

- [ ] **Step 1: Add DaemonAutoStart and DaemonAutoStop subs to Main.bas**

Add after `ExecuteMacDaemon`:

```vba
Function GetMacInterpreter() As String
    ' Shared helper: read interpreter from config (same logic as RunPython lines 47-55).
    GetMacInterpreter = GetConfig("INTERPRETER_MAC", "")
    If GetMacInterpreter = "" Then
        GetMacInterpreter = GetConfig("INTERPRETER", "python")
    End If
End Function

Function GetMacPythonPath() As String
    ' Shared helper: build PYTHONPATH string (same logic as RunPython lines 73-97).
    ' Extracted to avoid duplication between RunPython, DaemonAutoStart, etc.
    Dim AddExcelDir As String, ActiveFullName As String, ThisFullName As String

    AddExcelDir = GetConfig("ADD_WORKBOOK_TO_PYTHONPATH", "true")
    #If Mac Then
        If InStr(ActiveWorkbook.FullName, "://") = 0 Then
            ActiveFullName = ToPosixPath(ActiveWorkbook.FullName)
            ThisFullName = ToPosixPath(ThisWorkbook.FullName)
        Else
            ActiveFullName = ActiveWorkbook.FullName
            ThisFullName = ThisWorkbook.FullName
        End If
        GetMacPythonPath = AddExcelDir & ";" & ActiveFullName & ";" & ThisFullName & ";" & GetConfig("ONEDRIVE_CONSUMER_MAC") & ";" & GetConfig("ONEDRIVE_COMMERCIAL_MAC") & ";" & GetConfig("SHAREPOINT_MAC") & ";" & GetConfig("PYTHONPATH")
    #End If
End Function

Sub DaemonAutoStart()
    ' Call from Workbook_Open to auto-start daemon if DAEMON=1.
    ' Idempotent — safe to call multiple times.
    #If Mac Then
    Dim useDaemon As String

    useDaemon = GetConfig("DAEMON", "0")
    If useDaemon <> "1" And UCase(useDaemon) <> "TRUE" Then
        Exit Sub
    End If

    LaunchDaemon GetMacInterpreter(), GetMacPythonPath()
    #End If
End Sub

Sub DaemonAutoStop()
    ' Call from Workbook_BeforeClose to gracefully shut down daemon.
    #If Mac Then
    Dim useDaemon As String, interpreter As String, SocketPath As String

    useDaemon = GetConfig("DAEMON", "0")
    If useDaemon <> "1" And UCase(useDaemon) <> "TRUE" Then
        Exit Sub
    End If

    interpreter = GetConfig("INTERPRETER_MAC", "")
    If interpreter = "" Then
        interpreter = GetConfig("INTERPRETER", "python")
    End If

    SocketPath = GetDaemonSocketPath(ActiveWorkbook.FullName)

    ' Send SHUTDOWN via daemon_client — fire and forget
    Dim ParameterString As String
    ParameterString = interpreter
    ParameterString = ParameterString + "|" + SocketPath
    ParameterString = ParameterString + "|SHUTDOWN"
    ParameterString = ParameterString + "|2"  ' 2 second timeout

    On Error Resume Next
        AppleScriptTask "xlwings-" & XLWINGS_VERSION & ".applescript", "DaemonExecHandler", ParameterString
    On Error GoTo 0
    #End If
End Sub
```

- [ ] **Step 2: Verify VBA compiles**

This requires manual check in Excel VBA editor, or trust the syntax. The VBA follows existing patterns from `Main.bas`.

- [ ] **Step 3: Commit**

```bash
git add xlwings/addin/Main.bas
git commit -m "feat(daemon): add DaemonAutoStart/Stop subs for workbook lifecycle"
```

---

### Task 11: Integration smoke test script

**Files:**
- Create: `tests/test_daemon_integration.py`

This is a script for manual integration testing on a Mac with Excel. It verifies the full daemon lifecycle without Excel by simulating the VBA→AppleScript→daemon flow purely in Python.

- [ ] **Step 1: Write the integration test**

```python
# tests/test_daemon_integration.py
"""
Integration test for the full daemon lifecycle.
Runs without Excel — simulates the VBA→daemon flow in pure Python.

Usage: python -m pytest tests/test_daemon_integration.py -v -s
"""

import os
import subprocess
import sys
import tempfile
import time
import unittest


class TestDaemonIntegration(unittest.TestCase):
    """End-to-end test: start daemon, send EXEC, verify reload, shutdown."""

    def setUp(self):
        self.workbook_name = "IntegrationTest.xlsm"
        # Use explicit test-specific paths (no hash computation needed —
        # in production VBA computes the hash and passes paths to daemon via CLI args)
        self.socket_path = os.path.join(
            tempfile.gettempdir(),
            f"xlwings-daemon-integration-{os.getpid()}.sock",
        )
        self.pid_file_path = os.path.join(
            tempfile.gettempdir(),
            f"xlwings-daemon-integration-{os.getpid()}.pid",
        )
        self.module_dir = tempfile.mkdtemp()
        self.result_path = os.path.join(self.module_dir, "result.txt")

        # Create user module
        self.module_path = os.path.join(self.module_dir, "integration_mod.py")
        with open(self.module_path, "w") as f:
            f.write(
                f'def main():\n    with open("{self.result_path}", "w") as f:\n        f.write("v1")\n'
            )

        # Cleanup any leftover state
        for path in (self.socket_path, self.pid_file_path):
            if os.path.exists(path):
                os.unlink(path)

    def tearDown(self):
        # Kill any leftover daemon
        if os.path.exists(self.pid_file_path):
            try:
                with open(self.pid_file_path) as f:
                    pid = int(f.read().strip())
                os.kill(pid, 9)
            except (ProcessLookupError, ValueError):
                pass

        for path in (self.socket_path, self.pid_file_path,
                      self.module_path, self.result_path):
            if os.path.exists(path):
                os.unlink(path)
        status_path = self.socket_path.replace(".sock", ".status")
        if os.path.exists(status_path):
            os.unlink(status_path)
        os.rmdir(self.module_dir)

    def test_full_lifecycle(self):
        """Start daemon, EXEC, verify reload, SHUTDOWN."""
        python = sys.executable

        # 1. Start daemon in background
        daemon_process = subprocess.Popen(
            [
                python, "-m", "xlwings.daemon", "start",
                "--workbook", self.workbook_name,
                "--socket", self.socket_path,
                "--pidfile", self.pid_file_path,
                "--pythonpath", self.module_dir,
                "--app", "/Applications/Microsoft Excel.app/Contents/MacOS",
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )

        # 2. Wait for daemon to be ready
        deadline = time.monotonic() + 10
        while time.monotonic() < deadline:
            status_path = self.socket_path.replace(".sock", ".status")
            if os.path.exists(status_path):
                with open(status_path) as f:
                    if f.read().strip() == "ready":
                        break
            time.sleep(0.1)
        else:
            self.fail("Daemon did not become ready within 10 seconds")

        # 3. EXEC via daemon_client
        from xlwings.daemon_client import send_command

        response = send_command(
            self.socket_path,
            "EXEC import integration_mod; integration_mod.main()",
        )
        self.assertEqual(response, "OK")
        with open(self.result_path) as f:
            self.assertEqual(f.read(), "v1")

        # 4. Update module and verify reload
        with open(self.module_path, "w") as f:
            f.write(
                f'def main():\n    with open("{self.result_path}", "w") as f:\n        f.write("v2")\n'
            )

        response = send_command(
            self.socket_path,
            "EXEC import integration_mod; integration_mod.main()",
        )
        self.assertEqual(response, "OK")
        with open(self.result_path) as f:
            self.assertEqual(f.read(), "v2")

        # 5. PING
        response = send_command(self.socket_path, "PING")
        self.assertEqual(response, "PONG")

        # 6. SHUTDOWN
        response = send_command(self.socket_path, "SHUTDOWN")
        self.assertEqual(response, "OK")

        daemon_process.wait(timeout=5)
        self.assertEqual(daemon_process.returncode, 0)
```

- [ ] **Step 2: Run test**

Run: `python -m pytest tests/test_daemon_integration.py -v -s`
Expected: PASS (all 6 lifecycle steps succeed)

- [ ] **Step 3: Commit**

```bash
git add tests/test_daemon_integration.py
git commit -m "test(daemon): add end-to-end integration test for daemon lifecycle"
```

---

## Summary of all files

| File | Action | Task |
|------|--------|------|
| `xlwings/daemon.py` | Create | Tasks 1-5, 9 |
| `xlwings/daemon_client.py` | Create | Task 6 |
| `xlwings/addin/Main.bas` | Modify | Tasks 8, 10 |
| `xlwings/xlwings-dev.applescript` | Modify | Task 7 |
| `tests/test_daemon.py` | Create | Tasks 1-4, 9 |
| `tests/test_daemon_client.py` | Create | Task 6 |
| `tests/test_daemon_integration.py` | Create | Task 11 |

## Dependencies between tasks

```
Task 1 (socket server) ─┬─> Task 2 (EXEC) ──> Task 5 (CLI)
                         ├─> Task 3 (PID)
                         └─> Task 4 (health-check)
Task 5 (CLI) + Task 6 (client) ──> Task 7 (AppleScript) ──> Task 8 (VBA RunPython)
Task 8 ──> Task 10 (lifecycle hooks)
Task 9 (status file) is independent after Task 1
Task 11 (integration test) depends on Tasks 1-6, 9
```

Tasks 1→2→3→4 are sequential (each builds on the daemon server).
Tasks 5 and 9 can run in parallel after Task 4.
Task 6 can run in parallel with Tasks 3-5.
Tasks 7→8→10 are sequential (VBA/AppleScript integration).
Task 11 runs last as the integration test.

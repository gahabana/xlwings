# xlwings/daemon.py
"""
xlwings daemon — persistent Python process for fast RunPython execution on macOS.

Listens on a Unix domain socket and executes commands sent from Excel via VBA/AppleScript.
Eliminates 2-3s interpreter startup latency by keeping Python warm between calls.
"""

import fcntl
import importlib
import logging
import os
import socket
import sys
import traceback

logger = logging.getLogger(__name__)


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
        self._loaded_modules = {}  # module_name -> True (tracks which modules to reload)

    def serve(self):
        """Main loop: bind socket, accept connections, dispatch commands."""
        self._write_pid_file()
        self._running = True

        self._server_socket = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
        if os.path.exists(self.socket_path):
            os.unlink(self.socket_path)
        self._server_socket.bind(self.socket_path)
        os.chmod(self.socket_path, 0o600)
        self._server_socket.listen(1)
        self._server_socket.settimeout(1.0)

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

            # Remove previously-imported user modules so import picks up changes
            for module_name in list(self._loaded_modules):
                if module_name in sys.modules:
                    module = sys.modules[module_name]
                    # Delete cached .pyc so the fresh import reads the .py source
                    if hasattr(module, "__file__") and module.__file__:
                        try:
                            pyc_path = importlib.util.cache_from_source(
                                module.__file__
                            )
                            if os.path.exists(pyc_path):
                                os.unlink(pyc_path)
                        except (NotImplementedError, ValueError):
                            pass
                    del sys.modules[module_name]
            importlib.invalidate_caches()

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

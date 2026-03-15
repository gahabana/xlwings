# xlwings/daemon.py
"""
xlwings daemon — persistent Python process for fast RunPython execution on macOS.

Listens on a Unix domain socket and executes commands sent from Excel via VBA/AppleScript.
Eliminates 2-3s interpreter startup latency by keeping Python warm between calls.
"""

import argparse
import fcntl
import importlib
import logging
import os
import signal
import socket
import subprocess
import sys
import threading
import time
import traceback

logger = logging.getLogger(__name__)


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


class DaemonServer:
    """Unix domain socket server that accepts commands from Excel/VBA."""

    def __init__(
        self,
        socket_path,
        pid_file_path,
        workbook_name,
        pythonpath,
        app_path,
        health_check_interval=0,
    ):
        self.socket_path = socket_path
        self.pid_file_path = pid_file_path
        self.workbook_name = workbook_name
        self.pythonpath = pythonpath
        self.app_path = app_path
        self.health_check_interval = health_check_interval
        self.status_file_path = self.socket_path.replace(".sock", ".status")
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

        # Signal readiness
        with open(self.status_file_path, "w") as f:
            f.write("ready")

        # Start health-check thread if interval is set
        if self.health_check_interval > 0:
            health_thread = threading.Thread(
                target=self._health_check_loop, daemon=True
            )
            health_thread.start()

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

            # Use xlwings' own prepare_sys_path to set up sys.path correctly.
            # This handles extracting parent directories from file paths,
            # OneDrive/SharePoint URL resolution, etc.
            if self.pythonpath:
                from xlwings.utils import prepare_sys_path

                prepare_sys_path(self.pythonpath)

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
        if os.path.exists(self.status_file_path):
            os.unlink(self.status_file_path)

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

    def shutdown(self):
        """Signal the server to stop (called from health-check thread or signal handler)."""
        self._running = False


def check_stale_daemon(socket_path, pid_file_path):
    """Check if an existing daemon is alive. Returns True if alive, False if stale/dead.

    If a stale daemon is detected (PID file exists but process is dead or
    socket doesn't respond to PING), cleans up the stale PID/socket files.
    """
    if not os.path.exists(socket_path):
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
            pass
        os.unlink(pid_file_path)

    if os.path.exists(socket_path):
        os.unlink(socket_path)

    # Clean up stale status file too
    status_file_path = socket_path.replace(".sock", ".status")
    if os.path.exists(status_file_path):
        os.unlink(status_file_path)

    return False


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
    start_parser.add_argument("--health-check-interval", type=float, default=2.5,
                              help="Health-check interval in seconds (0 to disable)")

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
            health_check_interval=args.health_check_interval,
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

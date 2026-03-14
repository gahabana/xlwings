# xlwings/daemon.py
"""
xlwings daemon — persistent Python process for fast RunPython execution on macOS.

Listens on a Unix domain socket and executes commands sent from Excel via VBA/AppleScript.
Eliminates 2-3s interpreter startup latency by keeping Python warm between calls.
"""

import fcntl
import logging
import os
import socket

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

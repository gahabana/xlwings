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

        response = send_command("/tmp/nonexistent-socket.sock", "PING",
                                timeout=1)
        self.assertTrue(response.startswith("ERROR:"))

    def test_client_timeout_returns_timeout_error(self):
        """Client should return a specific timeout error for fallback detection."""
        from xlwings.daemon_client import send_command

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

        def delayed_start():
            time.sleep(0.5)
            server.serve()

        server_thread = threading.Thread(target=delayed_start, daemon=True)
        server_thread.start()

        response = send_command(self.socket_path, "PING",
                                timeout=5, retry_interval=0.2)
        self.assertEqual(response, "PONG")

        send_command(self.socket_path, "SHUTDOWN")
        server_thread.join(timeout=3)

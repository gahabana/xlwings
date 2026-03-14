# tests/test_daemon.py
import os
import shutil
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
        time.sleep(0.2)

        response = self._send_command("PING")
        self.assertEqual(response, "PONG")

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

    def test_exec_runs_python_code(self):
        """EXEC should execute Python code and return OK on success."""
        from xlwings.daemon import DaemonServer

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

        shutil.rmtree(module_dir)

    def test_exec_reloads_module_on_each_call(self):
        """EXEC should reload the user module so edits are picked up."""
        from xlwings.daemon import DaemonServer

        module_dir = tempfile.mkdtemp()
        module_path = os.path.join(module_dir, "reloadable.py")
        result_path = os.path.join(module_dir, "result.txt")

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

        response = self._send_command("EXEC import reloadable; reloadable.main()")
        self.assertEqual(response, "OK")
        with open(result_path) as f:
            self.assertEqual(f.read(), "v1")

        with open(module_path, "w") as f:
            f.write(
                f'def main():\n    with open("{result_path}", "w") as f:\n        f.write("v2")\n'
            )

        response = self._send_command("EXEC import reloadable; reloadable.main()")
        self.assertEqual(response, "OK")
        with open(result_path) as f:
            self.assertEqual(f.read(), "v2")

        self._send_command("SHUTDOWN")
        server_thread.join(timeout=3)

        shutil.rmtree(module_dir)

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

        response = self._send_command("EXEC raise ValueError('test error')")
        self.assertTrue(response.startswith("ERROR:"))
        self.assertIn("test error", response)

        response = self._send_command("PING")
        self.assertEqual(response, "PONG")

        self._send_command("SHUTDOWN")
        server_thread.join(timeout=3)


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

        with open(self.pid_file_path, "w") as f:
            f.write("999999999")

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

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
        escaped_result_path = self.result_path.replace("\\", "\\\\")
        with open(self.module_path, "w") as f:
            f.write(
                f'def main():\n    with open("{escaped_result_path}", "w") as f:\n        f.write("v1")\n'
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
        import shutil
        shutil.rmtree(self.module_dir, ignore_errors=True)

    def test_full_lifecycle(self):
        """Start daemon, EXEC, verify reload, SHUTDOWN."""
        python_executable = sys.executable

        # 1. Start daemon in background
        daemon_process = subprocess.Popen(
            [
                python_executable, "-m", "xlwings.daemon", "start",
                "--workbook", self.workbook_name,
                "--socket", self.socket_path,
                "--pidfile", self.pid_file_path,
                "--pythonpath", self.module_dir,
                "--app", "/Applications/Microsoft Excel.app/Contents/MacOS",
                "--health-check-interval", "0",
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
            stdout_output = daemon_process.stdout.read().decode() if daemon_process.stdout else ""
            stderr_output = daemon_process.stderr.read().decode() if daemon_process.stderr else ""
            self.fail(
                f"Daemon did not become ready within 10 seconds.\n"
                f"stdout: {stdout_output}\nstderr: {stderr_output}"
            )

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
        escaped_result_path = self.result_path.replace("\\", "\\\\")
        with open(self.module_path, "w") as f:
            f.write(
                f'def main():\n    with open("{escaped_result_path}", "w") as f:\n        f.write("v2")\n'
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

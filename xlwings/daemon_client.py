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

    if response.startswith("ERROR:"):
        sys.exit(1)


if __name__ == "__main__":
    main()

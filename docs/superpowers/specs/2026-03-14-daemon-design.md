# xlwings Daemon — Design Spec

## Problem

Every click of "Run" in Excel spawns a new Python process, which takes 2-3 seconds to start the interpreter, import xlwings/numpy/pandas, and import the user's module. For interactive workflows (e.g., Monte Carlo simulations), this startup latency dominates the total execution time.

## Solution

A persistent background Python daemon that starts when an xlwings-using workbook opens and stays running until the workbook closes. The daemon pre-imports all heavy modules so that subsequent "Run" clicks execute in ~50-100ms instead of 2-3 seconds.

## Scope

### In scope

- macOS implementation (AppleScript + Unix domain sockets)
- Per-workbook daemon processes with per-workbook venvs (uv-managed)
- Automatic lifecycle: start on workbook open, stop on workbook close
- Orphan cleanup for crash/force-quit scenarios
- Module reloading on each execution (always run latest saved code)
- Fallback to normal process spawn if daemon is unreachable

### Out of scope

- Windows implementation (different IPC mechanism needed — future work)
- UDF server mode (separate feature)
- Multi-workbook shared daemons

## Architecture

### Three components

1. **Daemon process** (`xlwings/daemon.py`) — Python script that:
   - Imports xlwings, numpy, pandas, and the user's module
   - Listens on a Unix domain socket
   - Writes a PID file for orphan detection
   - Runs a health-check thread monitoring Excel's process
   - Executes commands received over the socket
   - Reloads user module on every EXEC (fresh code each run)

2. **VBA modifications** (in `xlwings.xlam` addin):
   - `Workbook_Open`: check xlwings config for `DAEMON=1` setting, launch daemon in background if enabled
   - `Workbook_BeforeClose`: send shutdown command to daemon
   - Modified `RunPython` / `ExecuteMac`: connect to daemon socket instead of spawning process

3. **Stale process cleanup** — on daemon start, PING the existing socket first. If alive, reuse it (don't kill). If dead/unresponsive, acquire `fcntl.flock()` on PID file, kill stale process, start fresh

### File locations

| Artifact | Path |
|----------|------|
| Socket | `$TMPDIR/xlwings-daemon-{workbook_hash}.sock` (per-user, `chmod 0600`) |
| PID file | `$TMPDIR/xlwings-daemon-{workbook_hash}.pid` (locked with `fcntl.flock()`) |
| Daemon module | `xlwings/daemon.py` |
| VBA addin | `xlwings/addin/Main.bas` (modified) |
| AppleScript | `xlwings/xlwings-dev.applescript` (modified) |

The `{workbook_hash}` is derived from the workbook's full path to ensure uniqueness.

## Daemon Lifecycle

### Startup (triggered by Workbook_Open)

```
Excel opens PWandRC5.xlsm
  -> Workbook_Open fires in xlwings addin
  -> VBA checks xlwings config for DAEMON setting (not code scanning — explicit opt-in)
  -> DAEMON=1 -> reads Interpreter path from xlwings ribbon config
  -> AppleScript: do shell script "/path/to/uv/python -m xlwings.daemon start
      --workbook PWandRC5.xlsm --socket /tmp/xlwings-daemon-{hash}.sock &"
  -> Returns immediately (non-blocking, workbook opens at normal speed)
```

### Daemon boot (background, ~2-3 seconds)

```
1. Check for existing daemon -> PING socket -> if alive, exit (reuse existing)
2. Acquire fcntl.flock() on PID file (prevents race on rapid close-reopen)
3. Kill stale daemon if PID file exists but PING failed
4. Write own PID to locked PID file
5. Create Unix domain socket (chmod 0600 for security)
6. import xlwings (heavy import)
7. import numpy, pandas (if used)
8. import user_module (parsed from workbook config)
9. Write status file ($TMPDIR/xlwings-daemon-{hash}.status) with "ready" or "error: ..."
10. Start health-check thread (monitors Excel AND workbook, every 2-3s)
11. Accept connections on socket (signals "ready")
```

### Run command (user clicks "Run")

```
VBA RunPython("import PWandRC5; PWandRC5.main()")
  -> Try connect to /tmp/xlwings-daemon-{hash}.sock
  -> If connected: send EXEC command, wait for response
  -> If connection refused (daemon still booting): retry every 200ms, up to 10s
  -> If timeout after 10s: fall back to normal process spawn
```

The 10-second timeout with fallback ensures the user is **never stuck** — worst case they get the old 2-3 second spawn behavior.

### Shutdown (3 paths)

| Trigger | Mechanism |
|---------|-----------|
| Workbook closed normally | VBA `Workbook_BeforeClose` sends SHUTDOWN via socket |
| Excel crashes / force-quit | Health-check thread detects Excel PID gone, daemon exits |
| Stale from previous session | Next daemon start kills old PID before starting |

## Communication Protocol

Newline-delimited text over Unix domain socket. No HTTP, no JSON — minimal overhead.

### Messages (VBA -> Daemon)

```
EXEC import PWandRC5; PWandRC5.main()
PING
SHUTDOWN
```

### Responses (Daemon -> VBA)

```
OK
PONG
ERROR: Traceback (most recent call last):
  File "PWandRC5.py", line 42, in main
    ...
```

### Transport

VBA talks to the socket via AppleScript running a small Python helper (using the same interpreter as the daemon):

```bash
/path/to/python -c "
import socket, sys
s = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
s.connect('$TMPDIR/xlwings-daemon-{hash}.sock')
s.sendall(b'EXEC import PWandRC5; PWandRC5.main()\n')
response = b''
while True:
    chunk = s.recv(4096)
    if not chunk:
        break
    response += chunk
s.close()
print(response.decode(), end='')
"
```

This avoids the `nc -U` reliability issue (where `nc` may exit before reading the full response). The daemon closes the connection after sending the response, so EOF is the delimiter. This helper is a few lines and starts near-instantly since the interpreter is already cached by macOS.

Error responses are displayed in the same error dialog xlwings already uses for Python exceptions.

## Module Reloading

On every `EXEC` command, the daemon calls `importlib.reload(user_module)` before executing:

- Edit Python script -> click Run -> daemon picks up latest code (~5ms reload)
- If reload fails (syntax error), daemon returns the error but stays alive
- xlwings/numpy/pandas remain cached (the expensive imports that justify the daemon)

This matches the original behavior where each process spawn naturally gets fresh code via a clean `import`.

## Health-Check Thread

A background thread in the daemon that runs every 2-3 seconds, checking two conditions:

```python
while running:
    if not is_excel_running():
        cleanup_and_exit()
    if not is_workbook_open(workbook_name):
        cleanup_and_exit()
    time.sleep(2.5)
```

- **Excel detection:** `subprocess.check_output(["pgrep", "-x", "Microsoft Excel"])` — same approach xlwings uses in `_xlmac.py`
- **Workbook detection:** `osascript -e 'tell application "Microsoft Excel" to get name of every workbook'` — checks if the specific workbook is still open. This handles the case where the user closes the workbook but keeps Excel running.

## Environment Forwarding

The daemon must receive the same environment context that the current `ExecuteMac` passes (see `Main.bas` lines 164-169):

**At daemon startup (via command-line args):**
- `--pythonpath` — semicolon-delimited paths (workbook directory, etc.)
- `--workbook` — workbook name (for `--wb` arg and health-check)
- `--app` — Excel application path
- `--socket` — socket path

**At EXEC time (forwarded with each command):**
- The daemon sets `sys.argv` to match what `xlwings.utils.prepare_sys_path()` expects: `['--wb=<name>', '--from_xl=1', '--app=<path>']`
- This ensures the user's script sees the same environment as a normal RunPython call

## Files modified

| File | Change |
|------|--------|
| `xlwings/daemon.py` | **New file** — daemon server, socket handler, health-check, module reloader |
| `xlwings/addin/Main.bas` | Add Workbook_Open scan + daemon launch, modify ExecuteMac for socket path |
| `xlwings/xlwings-dev.applescript` | Add daemon launch and socket communication handlers |
| `tests/test_daemon.py` | **New file** — unit tests |

## Testing strategy

### Unit tests (no Excel required)

- Daemon socket lifecycle: start, accept connection, execute command, shutdown
- PID file management: create, detect stale, kill stale, clean up
- Health-check thread: mock Excel PID, verify self-termination when PID disappears
- Module reload: verify fresh code is picked up on each EXEC
- Protocol parsing: EXEC, PING, SHUTDOWN messages and responses
- Timeout/fallback: verify fallback to process spawn when daemon unreachable

### Manual integration test (requires Excel on Mac)

- Open xlwings workbook -> verify daemon process appears in Activity Monitor
- Click Run -> verify fast execution
- Edit Python script -> click Run -> verify new code runs
- Close workbook -> verify daemon process disappears
- Force-quit Excel -> verify daemon self-terminates within ~5 seconds

## Rollback

The daemon is explicitly opt-in via the `DAEMON` config setting. If anything goes wrong:

- Daemon timeout (10s) automatically falls back to normal process spawn
- Remove `DAEMON=1` from xlwings config or set `XLWINGS_DAEMON=0` env var to disable
- Status file (`$TMPDIR/xlwings-daemon-{hash}.status`) records boot errors for diagnosis
- Reverting the VBA changes restores original behavior completely

#!/bin/bash
#
# xlwings daemon wrapper — drop-in replacement for the Python interpreter.
#
# Set this as your Interpreter_Mac in xlwings ribbon instead of the path to python.
# It will:
#   1. Check if a daemon is running for the current workbook
#   2. If yes → send the command to the daemon (~50-100ms)
#   3. If no → start a daemon in background, run normally this time (~2-3s)
#
# Next click after daemon starts will be fast.
#
# To stop the daemon: close the workbook or kill it manually.

# Resolve the real Python interpreter from the venv
# This script should be symlinked/copied next to the venv's python,
# or XLWINGS_DAEMON_PYTHON can be set to the interpreter path.
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

if [ -n "$XLWINGS_DAEMON_PYTHON" ]; then
    PYTHON="$XLWINGS_DAEMON_PYTHON"
elif [ -x "$SCRIPT_DIR/python" ]; then
    PYTHON="$SCRIPT_DIR/python"
elif [ -x "$SCRIPT_DIR/python3" ]; then
    PYTHON="$SCRIPT_DIR/python3"
else
    PYTHON="python3"
fi

# Parse arguments from xlwings AppleScript:
# interpreter -B -u -W ignore -c "import xlwings.utils;xlwings.utils.prepare_sys_path('PYTHONPATH');USER_CODE" "--wb=Name" "--from_xl=1" "--app=Path"
WORKBOOK_NAME=""
APP_PATH=""
FULL_COMMAND=""

for arg in "$@"; do
    case "$arg" in
        --wb=*) WORKBOOK_NAME="${arg#--wb=}" ;;
        --app=*) APP_PATH="${arg#--app=}" ;;
    esac
done

# Get the -c argument (the full Python command)
NEXT_IS_COMMAND=false
for arg in "$@"; do
    if $NEXT_IS_COMMAND; then
        FULL_COMMAND="$arg"
        break
    fi
    if [ "$arg" = "-c" ]; then
        NEXT_IS_COMMAND=true
    fi
done

# If no workbook name found, just run Python normally
if [ -z "$WORKBOOK_NAME" ]; then
    exec "$PYTHON" "$@"
fi

# Determine socket path using md5 hash of workbook name
TMPDIR_RESOLVED="${TMPDIR:-/tmp/}"
SOCKET_HASH=$(printf '%s' "$WORKBOOK_NAME" | md5 | cut -c1-12)
SOCKET_PATH="${TMPDIR_RESOLVED}xlwings-daemon-${SOCKET_HASH}.sock"
PID_FILE="${TMPDIR_RESOLVED}xlwings-daemon-${SOCKET_HASH}.pid"

# Check if daemon is alive
if [ -S "$SOCKET_PATH" ]; then
    PONG=$("$PYTHON" -m xlwings.daemon_client "$SOCKET_PATH" "PING" --timeout 2 2>/dev/null)
    if [ "$PONG" = "PONG" ]; then
        # Daemon is alive — extract user code and send via daemon
        # Full command: "import xlwings.utils;xlwings.utils.prepare_sys_path('...');import foo; foo.main()"
        # User code is everything after prepare_sys_path(...);
        USER_CODE=$(printf '%s' "$FULL_COMMAND" | sed "s/.*prepare_sys_path([^)]*)[;] *//" )

        RESULT=$("$PYTHON" -m xlwings.daemon_client "$SOCKET_PATH" "EXEC $USER_CODE" --timeout 30 2>&1)

        if printf '%s' "$RESULT" | grep -q "^ERROR:"; then
            # Write error to log file for Excel to pick up
            LOG_FILE="$HOME/xlwings.log"
            printf '%s' "$RESULT" | sed 's/^ERROR: //' > "$LOG_FILE"
            exit 1
        fi
        exit 0
    fi
fi

# Daemon not running — start it in background, then run normally this time
# Extract PYTHONPATH from prepare_sys_path argument
PYTHONPATH_ARG=$(printf '%s' "$FULL_COMMAND" | sed "s/.*prepare_sys_path('\\([^']*\\)').*/\\1/")

"$PYTHON" -m xlwings.daemon start \
    --workbook "$WORKBOOK_NAME" \
    --socket "$SOCKET_PATH" \
    --pidfile "$PID_FILE" \
    --pythonpath "$PYTHONPATH_ARG" \
    --app "${APP_PATH:-/Applications/Microsoft Excel.app/Contents/MacOS}" \
    --health-check-interval 2.5 \
    > /dev/null 2>&1 &

# Run normally this time (daemon will be ready for next call)
exec "$PYTHON" "$@"

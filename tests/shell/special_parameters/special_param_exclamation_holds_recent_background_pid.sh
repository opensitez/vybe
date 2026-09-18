#!/usr/bin/env bash
# vybe-test: bash/special_parameters/special_param_exclamation_holds_recent_background_pid
# The $! parameter expands to the process ID of the most recently placed background job.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
sleep 0.05 &
bg_pid=$!
wait "$bg_pid"
[ "$bg_pid" -gt 0 ] 2>/dev/null || fail "\$! is not a valid background PID: got [$bg_pid]"
echo PASS
exit 0

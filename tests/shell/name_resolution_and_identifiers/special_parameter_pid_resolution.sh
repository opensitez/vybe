#!/usr/bin/env bash
# vybe-test: bash/name_resolution_and_identifiers/special_parameter_pid_resolution
# The special parameter identifier '$' resolves to the process ID of the invoking shell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
pid=$$
[ "$pid" -gt 0 ] || fail "shell pid: want > 0, got $pid"
echo PASS
exit 0

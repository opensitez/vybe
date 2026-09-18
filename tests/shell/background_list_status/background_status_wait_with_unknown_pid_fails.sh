#!/usr/bin/env bash
# vybe-test: bash/background_list_status/background_status_wait_with_unknown_pid_fails
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
wait 999999 2>/dev/null
got=$?
[ "$got" -eq 1 ] || fail "unknown pid should fail"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/background_list_execution/background_background_status_failure
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
( false ) &
pid=$!
wait "$pid"
st=$?
[ "$st" -eq 1 ] || fail "background failure should return 1"
rm -f /tmp/bash_bg_exec_fail_$PPID
echo PASS
exit 0

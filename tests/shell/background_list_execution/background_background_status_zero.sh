#!/usr/bin/env bash
# vybe-test: bash/background_list_execution/background_background_status_zero
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: > /tmp/bash_bg_exec_status_$PPID
( : ) &
pid=$!
wait "$pid"
st=$?
[ "$st" -eq 0 ] || fail "background success should return 0"
rm -f /tmp/bash_bg_exec_status_$PPID
echo PASS
exit 0

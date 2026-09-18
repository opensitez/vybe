#!/usr/bin/env bash
# vybe-test: bash/background_list_execution/background_background_pid_is_positive_integer
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: &
p=$!
wait "$p"
[ "$p" -gt 0 ] || fail "pid should be positive"
rm -f /tmp/bash_bg_exec_pid_$PPID
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/background_list_status/status_wait_respects_brittle_pid_cleanup
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: &
p=$!
wait "$p"
x=$?
wait "$p"
y=$?
[ "$x" -eq 0 ] || fail "first wait 0"
[ "$y" -ne 0 ] || fail "second wait should fail after reaping"
echo PASS
exit 0

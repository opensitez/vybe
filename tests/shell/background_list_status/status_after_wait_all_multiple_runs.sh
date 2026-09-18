#!/usr/bin/env bash
# vybe-test: bash/background_list_status/status_after_wait_all_multiple_runs
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: &
: &
wait
a=$?
wait
[ "$a" -eq 0 ] || fail "first wait should be 0"
[ "$?" -eq 0 ] || fail "second wait should also be 0"
echo PASS
exit 0

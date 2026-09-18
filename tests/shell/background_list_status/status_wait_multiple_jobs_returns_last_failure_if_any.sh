#!/usr/bin/env bash
# vybe-test: bash/background_list_status/status_wait_multiple_jobs_returns_last_failure_if_any
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: &
p1=$!
false &
p2=$!
wait "$p1"
[ "$?" -eq 0 ] || fail "first success must be 0"
wait "$p2"
[ "$?" -eq 1 ] || fail "second false must be 1"
echo PASS
exit 0

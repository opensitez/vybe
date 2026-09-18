#!/usr/bin/env bash
# vybe-test: bash/child_status_collection/wait_collects_last_failing_status
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: &
p1=$!
false &
p2=$!
wait "$p1"
wait "$p2"
[ $? -eq 1 ] || fail "last failure should be visible"
echo PASS
exit 0

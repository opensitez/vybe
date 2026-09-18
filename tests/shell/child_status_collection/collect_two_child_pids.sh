#!/usr/bin/env bash
# vybe-test: bash/child_status_collection/collect_two_child_pids
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: &
p1=$!
: &
p2=$!
[ "$p1" != "$p2" ] || fail "different jobs"
wait "$p1"
wait "$p2"
echo PASS
exit 0

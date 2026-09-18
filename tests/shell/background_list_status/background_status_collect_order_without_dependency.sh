#!/usr/bin/env bash
# vybe-test: bash/background_list_status/background_status_collect_order_without_dependency
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: &
p1=$!
( : ) &
p2=$!
wait "$p1"
st1=$?
wait "$p2"
st2=$?
[ "$st1" -eq 0 ] || fail "first job status"
[ "$st2" -eq 0 ] || fail "second job status"
[ "$p1" != "$p2" ] || fail "pids should differ"
echo PASS
exit 0

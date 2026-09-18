#!/usr/bin/env bash
# vybe-test: bash/background_list_status/background_wait_one_by_one
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: &
p1=$!
false &
p2=$!
wait "$p1"; s1=$?
wait "$p2"; s2=$?
[ "$s1" -eq 0 ] || fail "first status"
[ "$s2" -eq 1 ] || fail "second status"
echo PASS
exit 0

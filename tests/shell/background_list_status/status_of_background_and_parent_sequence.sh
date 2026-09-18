#!/usr/bin/env bash
# vybe-test: bash/background_list_status/status_of_background_and_parent_sequence
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=0
false &
p=$!
wait "$p"
[ "$?" -eq 1 ] && a=1
[ "$a" -eq 1 ] || fail "parent should continue and track status"
echo PASS
exit 0

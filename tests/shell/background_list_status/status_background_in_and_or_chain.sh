#!/usr/bin/env bash
# vybe-test: bash/background_list_status/status_background_in_and_or_chain
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: && : &
p=$!
wait "$p"
[ "$?" -eq 0 ] || fail "and-list background should pass"
false || : &
p2=$!
wait "$p2"
[ "$?" -eq 0 ] || fail "or-list background should pass"
echo PASS
exit 0

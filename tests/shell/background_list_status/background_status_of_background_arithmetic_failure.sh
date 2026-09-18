#!/usr/bin/env bash
# vybe-test: bash/background_list_status/background_status_of_background_arithmetic_failure
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
( ((1/0)) ) &
p=$!
wait "$p"
[ "$?" -eq 1 ] || fail "background arithmetic errors should fail"
echo PASS
exit 0

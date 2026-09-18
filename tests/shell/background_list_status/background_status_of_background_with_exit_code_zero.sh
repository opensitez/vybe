#!/usr/bin/env bash
# vybe-test: bash/background_list_status/background_status_of_background_with_exit_code_zero
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
( exit 0 ) &
p=$!
wait "$p"
[ "$?" -eq 0 ] || fail "exit 0 expected"
echo PASS
exit 0

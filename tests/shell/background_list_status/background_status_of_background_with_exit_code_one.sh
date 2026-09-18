#!/usr/bin/env bash
# vybe-test: bash/background_list_status/background_status_of_background_with_exit_code_one
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
( exit 1 ) &
p=$!
wait "$p"
[ "$?" -eq 1 ] || fail "exit 1 expected"
echo PASS
exit 0

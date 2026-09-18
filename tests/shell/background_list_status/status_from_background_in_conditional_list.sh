#!/usr/bin/env bash
# vybe-test: bash/background_list_status/status_from_background_in_conditional_list
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
true && false &
p=$!
wait "$p"
[ "$?" -eq 1 ] || fail "list status should be false branch status"
echo PASS
exit 0

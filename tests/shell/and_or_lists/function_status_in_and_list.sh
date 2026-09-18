#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/function_status_in_and_list
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
pass() { :; }
fail_fn() { return 1; }
log=""
pass && log=pass
[ "$log" = "pass" ] || fail "passing function in and"
log=""
pass && fail_fn && log=bad
[ "$log" = "" ] || fail "failing function should stop chain"
echo PASS
exit 0

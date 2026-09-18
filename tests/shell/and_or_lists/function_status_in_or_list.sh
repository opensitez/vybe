#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/function_status_in_or_list
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
pass() { return 0; }
fail_fn() { return 1; }
log=""
pass || log=bad
[ "$log" = "" ] || fail "or list should skip on success"
log=""
fail_fn || log=from_or
[ "$log" = "from_or" ] || fail "or list should run after failure"
echo PASS
exit 0

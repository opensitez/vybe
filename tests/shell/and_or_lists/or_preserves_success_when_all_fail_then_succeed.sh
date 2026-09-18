#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/or_preserves_success_when_all_fail_then_succeed
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
# Bash `||` keeps running until first success; if one succeeds, list status is 0.
false || false || :
[ $? -eq 0 ] || fail "trailing : should make or-list succeed"
false || false || false
[ $? -eq 1 ] || fail "all false should yield failure"
echo PASS
exit 0

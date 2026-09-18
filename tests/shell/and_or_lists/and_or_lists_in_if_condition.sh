#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/and_or_lists_in_if_condition
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
branch=none
if : && :; then branch=and_ok; else branch=and_bad; fi
[ "$branch" = "and_ok" ] || fail "and-chain if condition should pass"
branch=none
if false || :; then branch=or_ok; else branch=or_bad; fi
[ "$branch" = "or_ok" ] || fail "or-chain if condition should pass"
echo PASS
exit 0

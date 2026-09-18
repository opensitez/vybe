#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/and_skips_rhs_when_lhs_fails
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
false && x=$((x + 1))
[ "$x" -eq 0 ] || fail "rhs must be skipped on lhs failure"
echo PASS
exit 0

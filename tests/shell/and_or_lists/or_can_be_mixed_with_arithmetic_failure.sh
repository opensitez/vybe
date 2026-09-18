#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/or_can_be_mixed_with_arithmetic_failure
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
((0)) || x=9
[ "$x" -eq 9 ] || fail "arithmetic failure should trigger rhs"
echo PASS
exit 0

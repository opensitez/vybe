#!/usr/bin/env bash
# vybe-test: bash/conditional_expression_syntax/cond_arithmetic_greater_or_equal_dash_ge
# The -ge operator tests numerical greater-than-or-equal-to comparisons.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ 20 -ge 10 ]] || fail "20 should be -ge 10"
[[ 10 -ge 10 ]] || fail "10 should be -ge 10"
[[ 9 -ge 10 ]] && fail "9 should not be -ge 10"
echo PASS
exit 0

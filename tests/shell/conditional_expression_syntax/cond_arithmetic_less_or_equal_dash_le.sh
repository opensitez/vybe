#!/usr/bin/env bash
# vybe-test: bash/conditional_expression_syntax/cond_arithmetic_less_or_equal_dash_le
# The -le operator tests numerical less-than-or-equal-to comparisons.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ 5 -le 10 ]] || fail "5 should be -le 10"
[[ 10 -le 10 ]] || fail "10 should be -le 10"
[[ 11 -le 10 ]] && fail "11 should not be -le 10"
echo PASS
exit 0

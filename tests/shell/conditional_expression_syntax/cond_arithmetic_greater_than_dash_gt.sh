#!/usr/bin/env bash
# vybe-test: bash/conditional_expression_syntax/cond_arithmetic_greater_than_dash_gt
# The -gt operator tests numerical strictly-greater-than comparisons.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ 20 -gt 10 ]] || fail "20 should be -gt 10"
[[ 10 -gt 20 ]] && fail "10 should not be -gt 20"
[[ 10 -gt 10 ]] && fail "10 should not be -gt 10"
echo PASS
exit 0

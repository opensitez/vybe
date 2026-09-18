#!/usr/bin/env bash
# vybe-test: bash/conditional_expression_syntax/cond_arithmetic_less_than_dash_lt
# The -lt operator tests numerical strictly-less-than comparisons.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ 5 -lt 10 ]] || fail "5 should be -lt 10"
[[ 10 -lt 5 ]] && fail "10 should not be -lt 5"
[[ 10 -lt 10 ]] && fail "10 should not be -lt 10"
echo PASS
exit 0

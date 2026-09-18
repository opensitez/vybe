#!/usr/bin/env bash
# vybe-test: bash/conditional_expression_syntax/cond_arithmetic_inequality_dash_ne
# The -ne operator tests numerical integer inequality.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ 10 -ne 20 ]] || fail "10 should not equal 20"
[[ 10 -ne 10 ]] && fail "10 -ne 10 should return false"
echo PASS
exit 0

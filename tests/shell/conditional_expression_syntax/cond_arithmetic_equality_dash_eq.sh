#!/usr/bin/env bash
# vybe-test: bash/conditional_expression_syntax/cond_arithmetic_equality_dash_eq
# The -eq operator tests numerical integer equality.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ 42 -eq 42 ]] || fail "42 should equal 42"
[[ 052 -eq 42 ]] || fail "octal 052 should equal decimal 42 in -eq"
[[ 42 -eq 43 ]] && fail "42 should not equal 43"
echo PASS
exit 0

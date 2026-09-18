#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/double_minus_is_decrement_not_double_negation
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=3
[ $((--x)) -eq 2 ] || fail "--x is pre-decrement, got $((--x))"
x=3
[ $((- -x)) -eq 3 ] || fail "- -x is double negation, got $((- -x))"
[ "$x" -eq 3 ] || fail "double negation must not modify x"
echo PASS
exit 0

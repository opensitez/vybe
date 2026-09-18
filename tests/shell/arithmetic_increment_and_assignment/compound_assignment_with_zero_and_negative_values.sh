#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/compound_assignment_with_zero_and_negative_values
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
[ $((x += -5)) -eq -5 ] || fail "x += -5 wrong"
[ "$x" -eq -5 ] || fail "x should be -5"
[ $((x -= -3)) -eq -2 ] || fail "x -= -3 wrong"
[ "$x" -eq -2 ] || fail "x should be -2"
[ $((x *= -1)) -eq 2 ] || fail "x *= -1 wrong"
[ "$x" -eq 2 ] || fail "x should be 2"
[ $((x %= 2)) -eq 0 ] || fail "x %= 2 wrong"
[ "$x" -eq 0 ] || fail "x should be 0"
echo PASS

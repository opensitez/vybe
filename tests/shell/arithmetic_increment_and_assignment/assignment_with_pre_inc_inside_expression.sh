#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/assignment_with_pre_inc_inside_expression
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=1
y=$((x + ++x))
[ "$y" -eq 3 ] || fail "x + ++x should be 3 got $y"
[ "$x" -eq 2 ] || fail "x should be 2"
z=$((++x * 3))
[ "$z" -eq 9 ] || fail "++x * 3 should be 9 got $z"
[ "$x" -eq 3 ] || fail "x should be 3"
echo PASS
exit 0

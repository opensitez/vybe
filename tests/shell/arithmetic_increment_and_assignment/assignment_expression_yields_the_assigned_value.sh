#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/assignment_expression_yields_the_assigned_value
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((x = 5)) -eq 5 ] || fail "value of assignment"
[ "$x" -eq 5 ] || fail "x not assigned"
[ $((y = x * 2)) -eq 10 ] && [ "$y" -eq 10 ] || fail "y=$y"
echo PASS
exit 0

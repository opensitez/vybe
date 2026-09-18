#!/usr/bin/env bash
# vybe-test: bash/scalar_variable_assignment/scalar_assignment_with_arithmetic_expansion
# The right-hand side of a scalar variable assignment can be an arithmetic expansion $(( ... )).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=15
b=25
result=$(( (a + b) * 2 ))
[ "$result" -eq 80 ] || fail "assignment with arithmetic expansion: want 80, got $result"
echo PASS
exit 0

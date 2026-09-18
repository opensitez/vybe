#!/usr/bin/env bash
# vybe-test: bash/scalar_variable_assignment/scalar_assignment_references_prior_assignment_same_line
# Assignments on the same line are evaluated left-to-right, making earlier assignments visible immediately.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=10 y=$(( x * 2 )) z=$(( y + 5 ))
[ "$x" -eq 10 ] || fail "x: want 10, got $x"
[ "$y" -eq 20 ] || fail "y: want 20, got $y"
[ "$z" -eq 25 ] || fail "z: want 25, got $z"
echo PASS
exit 0

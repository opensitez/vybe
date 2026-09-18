#!/usr/bin/env bash
# vybe-test: bash/integer_arithmetic/integer_arithmetic_assignment_expression_value
# Assignment inside arithmetic returns the assigned RHS value.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=1
result=$((x = 5 + 2))
[ "$result" -eq 7 ] || fail "result want 7 got $result"
[ "$x" -eq 7 ] || fail "x must become 7, got $x"
echo PASS
exit 0

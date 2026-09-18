#!/usr/bin/env bash
# vybe-test: bash/arithmetic_operator_precedence/comma_is_below_assignment
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
r=$((y = 1, 2))
[ "$r" -eq 2 ] || fail "result want 2 got $r"
[ "$y" -eq 1 ] || fail "y = 1 binds before the comma, y=$y"
echo PASS
exit 0

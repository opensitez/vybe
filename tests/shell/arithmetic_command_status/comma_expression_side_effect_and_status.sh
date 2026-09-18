#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/comma_expression_side_effect_and_status
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
(( x = 1, y = 2, x + y )); st=$?
[ "$st" -eq 0 ] || fail "1+2 should be success"
[ "$x" -eq 1 ] || fail "x should be 1"
[ "$y" -eq 2 ] || fail "y should be 2"
(( x = 0, y = 5, x / y )); st=$?
[ "$st" -eq 1 ] || fail "0 / 5 should be 0 and fail"
[ "$x" -eq 0 ] || fail "x should stay 0"
[ "$y" -eq 5 ] || fail "y should be 5"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/increment_in_arithmetic_expression_list
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=1
y=$((x += 1, x++, x))
[ "$y" -eq 3 ] || fail "list assignment/increment final value should be 3 got $y"
[ "$x" -eq 3 ] || fail "x should be 3"
x=4
y=$((x++, x += 2, x*2))
[ "$y" -eq 10 ] || fail "expect 10 got $y"
[ "$x" -eq 6 ] || fail "x should be 6"
echo PASS
exit 0

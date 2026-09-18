#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/increment_then_divide_rounds_toward_zero
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=-7
y=$((++x / 3))
[ "$x" -eq -6 ] || fail "x should be -6"
[ "$y" -eq -2 ] || fail "-6/3 = -2"
z=$((x-- / 2))
[ "$x" -eq -7 ] || fail "x should be restored to -7"
[ "$z" -eq -3 ] || fail "-6/2 = -3"
echo PASS
exit 0

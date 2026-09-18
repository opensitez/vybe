#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/compound_assignment_to_nested_expression
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=2
y=$(( x += 3 ))
[ "$y" -eq 5 ] || fail "x+=3 returns 5"
[ "$x" -eq 5 ] || fail "x should be 5"
x=10
y=$(( x >>= 1 ))
[ "$y" -eq 5 ] || fail "x>>=1 returns 5"
[ "$x" -eq 5 ] || fail "x should be 5"
x=10
y=$(( x ^= 3 ))
[ "$y" -eq 9 ] || fail "x^=3 returns 9"
[ "$x" -eq 9 ] || fail "x should be 9"
echo PASS
exit 0

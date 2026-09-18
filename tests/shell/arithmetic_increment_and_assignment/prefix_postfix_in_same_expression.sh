#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/prefix_postfix_in_same_expression
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=3
[ $((++x + x--)) -eq 7 ] || fail "++x + x-- expected 7 got $((++x + x--))"
[ "$x" -eq 3 ] || fail "x should restore to 3"
x=3
[ $((x++ + ++x)) -eq 8 ] || fail "x++ + ++x expected 8 got $((x++ + ++x))"
[ "$x" -eq 5 ] || fail "x final 5 got $x"
echo PASS
exit 0

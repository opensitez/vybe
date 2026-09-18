#!/usr/bin/env bash
# vybe-test: bash/arithmetic_operator_precedence/assignment_is_lowest_and_right_associative
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
r=$((a = b = 3))
[ "$r" -eq 3 ] && [ "$a" -eq 3 ] && [ "$b" -eq 3 ] || fail "chain: r=$r a=$a b=$b"
r=$((x = 1 + 2 * 3))
[ "$r" -eq 7 ] && [ "$x" -eq 7 ] || fail "whole expression assigned: r=$r x=$x"
echo PASS
exit 0

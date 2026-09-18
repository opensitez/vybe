#!/usr/bin/env bash
# vybe-test: bash/integer_arithmetic/operands_from_strings_are_trimmed_and_evaluated
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=" 42 "
[ $((x + 1)) -eq 43 ] || fail "surrounding blanks: got $((x + 1))"
y="1+2"
[ $((y * 2)) -eq 6 ] || fail "string value is an expression: got $((y * 2))"
[ $(( "3" + 1 )) -eq 4 ] || fail "quoted literal: got $(( "3" + 1 ))"
echo PASS
exit 0

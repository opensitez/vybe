#!/usr/bin/env bash
# vybe-test: bash/scalar_variable_assignment/scalar_assignment_identifier_with_digits
# Variable identifiers may contain digits as long as the first character is not a digit.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
var1="first"
v2b3="second"
_0="digit_after_underscore"
[ "$var1" = "first" ] || fail "var1 failed"
[ "$v2b3" = "second" ] || fail "v2b3 failed"
[ "$_0" = "digit_after_underscore" ] || fail "_0 failed"
echo PASS
exit 0

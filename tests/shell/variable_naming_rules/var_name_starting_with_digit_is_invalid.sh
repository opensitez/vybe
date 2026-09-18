#!/usr/bin/env bash
# vybe-test: bash/variable_naming_rules/var_name_starting_with_digit_is_invalid
# Variable identifiers starting with a digit are invalid and cannot be assigned.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval '1var=invalid' 2>/dev/null
st1=$?
[ "$st1" -ne 0 ] || fail "1var=invalid should fail"

eval '99_count=invalid' 2>/dev/null
st2=$?
[ "$st2" -ne 0 ] || fail "99_count=invalid should fail"
echo PASS
exit 0

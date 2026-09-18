#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/non_numeric_value_is_looked_up_as_a_name
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset abc
x=abc
[ $((x)) -eq 0 ] || fail "unset target is 0, got $((x))"
abc=7
[ $((x)) -eq 7 ] || fail "set target is used, got $((x))"
echo PASS
exit 0

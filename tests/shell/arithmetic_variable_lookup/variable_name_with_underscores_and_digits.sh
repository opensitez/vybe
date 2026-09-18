#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/variable_name_with_underscores_and_digits
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
var_1=9
v2=4
[ $((var_1 + v2)) -eq 13 ] || fail "expected 13 got $((var_1 + v2))"
echo PASS
exit 0

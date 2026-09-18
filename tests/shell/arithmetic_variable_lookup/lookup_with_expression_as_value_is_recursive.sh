#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/lookup_with_expression_as_value_is_recursive
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=3
b='a + 4'
c='b * 2'
[ $((c)) -eq 14 ] || fail "c expression chain expected 14 got $((c))"
echo PASS
exit 0

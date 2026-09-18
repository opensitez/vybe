#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/lookup_resolves_unset_to_zero_in_nested_expression
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset x y
x='y + 2'
[ $((x)) -eq 2 ] || fail "unset y should be 0"
[ $((x + 5)) -eq 7 ] || fail "nested expression with unset"
echo PASS
exit 0

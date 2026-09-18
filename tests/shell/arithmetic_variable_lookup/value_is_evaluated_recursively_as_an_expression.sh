#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/value_is_evaluated_recursively_as_an_expression
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
z=3; y='z * 2'; x='y + 1'
[ $((x)) -eq 7 ] || fail "want 7 got $((x))"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/scalar_variable_assignment/scalar_integer_attribute_append_addition
# When an integer attribute is set via 'declare -i', '+=' performs arithmetic addition instead of string concat.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -i count=100
count+=50
[ "$count" -eq 150 ] || fail "integer append addition: want 150, got $count"
count+=-20
[ "$count" -eq 130 ] || fail "integer append negative addition: want 130, got $count"
echo PASS
exit 0

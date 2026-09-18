#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/indirect_name_lookup_from_variable
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
base=5
name=base
[ $((name)) -eq 5 ] || fail "indirect by name variable"
name=base2
base2=12
[ $((name)) -eq 12 ] || fail "renamed indirection wrong"
echo PASS
exit 0

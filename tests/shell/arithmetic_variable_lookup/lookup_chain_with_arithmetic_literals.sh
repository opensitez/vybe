#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/lookup_chain_with_arithmetic_literals
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=6
b='a'
c='b'
[ $((c)) -eq 6 ] || fail "nested lookup chain to literal name"
[ $(( $c )) -eq 6 ] || fail "dollar-indirection to nested name"
echo PASS
exit 0

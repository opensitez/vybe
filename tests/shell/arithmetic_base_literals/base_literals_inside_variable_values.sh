#!/usr/bin/env bash
# vybe-test: bash/arithmetic_base_literals/base_literals_inside_variable_values
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0x10; y=2#11; z=010
[ $((x)) -eq 16 ] || fail "hex in variable got $((x))"
[ $((y)) -eq 3 ] || fail "base# in variable got $((y))"
[ $((z)) -eq 8 ] || fail "octal in variable got $((z))"
echo PASS
exit 0

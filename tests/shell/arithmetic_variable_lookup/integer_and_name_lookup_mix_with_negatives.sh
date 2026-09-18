#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/integer_and_name_lookup_mix_with_negatives
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=-3
b='a'
[ $((b)) -eq -3 ] || fail "b points at a"
[ $((b + 10)) -eq 7 ] || fail "b+10"
x='b * -2'
[ $((x)) -eq 6 ] || fail "negated chain"
echo PASS
exit 0

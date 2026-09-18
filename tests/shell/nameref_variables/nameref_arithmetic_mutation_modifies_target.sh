#!/usr/bin/env bash
# vybe-test: bash/nameref_variables/nameref_arithmetic_mutation_modifies_target
# Arithmetic operations that mutate a nameref like '(( ref++ ))' mutate the underlying target variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
raw_count=20
declare -n ref=raw_count
(( ref += 5 ))
[ "$raw_count" -eq 25 ] || fail "arithmetic addition failed: want 25, got $raw_count"
(( ref++ ))
[ "$raw_count" -eq 26 ] || fail "arithmetic increment failed: want 26, got $raw_count"
echo PASS
exit 0

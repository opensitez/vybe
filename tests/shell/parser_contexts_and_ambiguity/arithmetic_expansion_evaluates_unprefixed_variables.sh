#!/usr/bin/env bash
# vybe-test: bash/parser_contexts_and_ambiguity/arithmetic_expansion_evaluates_unprefixed_variables
# Inside $(( ... )), variables are evaluated as numeric values without requiring the '$' prefix.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
val_a=15
val_b=25
calc=$(( val_a * 2 + val_b ))
[ "$calc" -eq 55 ] || fail "unprefixed arithmetic expansion: want 55, got $calc"
echo PASS
exit 0

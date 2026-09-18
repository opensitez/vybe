#!/usr/bin/env bash
# vybe-test: bash/variable_naming_rules/var_name_single_character_identifier
# Single-letter identifiers such as 'a', 'i', 'Z' are valid variables.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=1; b=2; Z=26
sum=$(( a + b + Z ))
[ "$sum" -eq 29 ] || fail "single-character identifiers failed: want 29, got $sum"
echo PASS
exit 0

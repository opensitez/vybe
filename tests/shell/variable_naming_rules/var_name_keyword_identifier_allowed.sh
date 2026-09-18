#!/usr/bin/env bash
# vybe-test: bash/variable_naming_rules/var_name_keyword_identifier_allowed
# Shell reserved words (such as 'if', 'then', 'while', 'do') may be used as variable names.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if=100
then=200
while=300
do=400
total=$(( if + then + while + do ))
[ "$total" -eq 1000 ] || fail "reserved words as variable names failed: want 1000, got $total"
echo PASS
exit 0

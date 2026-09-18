#!/usr/bin/env bash
# vybe-test: bash/variable_naming_rules/var_name_numbers_after_letters_supported
# Digits are fully valid in variable identifiers as long as they follow at least one letter or underscore.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x1=10
y2=20
_3=30
sum=$(( x1 + y2 + _3 ))
[ "$sum" -eq 60 ] || fail "digits in variable identifiers failed: want 60, got $sum"
echo PASS
exit 0

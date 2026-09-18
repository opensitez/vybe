#!/usr/bin/env bash
# vybe-test: bash/variable_lookup_and_unset/var_lookup_in_arithmetic_without_dollar_prefix
# Variables referenced inside $(( ... )) are automatically evaluated as numbers without leading '$'.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
base_val=30
mult_val=4
computed=$(( base_val * mult_val + 5 ))
[ "$computed" -eq 125 ] || fail "arithmetic variable lookup without dollar: want 125, got $computed"
echo PASS
exit 0

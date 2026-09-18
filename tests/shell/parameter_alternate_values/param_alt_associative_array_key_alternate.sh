#!/usr/bin/env bash
# vybe-test: bash/parameter_alternate_values/param_alt_associative_array_key_alternate
# The alternate value expansion applies to specific associative array keys.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map=( [present]="data" )
res_pres="${map[present]:+key_present}"
res_abs="${map[absent]:+key_present}"
[ "$res_pres" = "key_present" ] || fail "present key alternate failed: got [$res_pres]"
[ -z "$res_abs" ] || fail "absent key alternate should be empty: got [$res_abs]"
echo PASS
exit 0

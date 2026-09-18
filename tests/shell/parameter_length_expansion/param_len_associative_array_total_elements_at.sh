#!/usr/bin/env bash
# vybe-test: bash/parameter_length_expansion/param_len_associative_array_total_elements_at
# The ${#map[@]} expansion returns the total number of entries in an associative array.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map=( [user]="admin" [role]="dev" [tier]="gold" )
[ "${#map[@]}" -eq 3 ] || fail "associative array size: want 3, got ${#map[@]}"
echo PASS
exit 0

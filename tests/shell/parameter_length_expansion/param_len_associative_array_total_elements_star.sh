#!/usr/bin/env bash
# vybe-test: bash/parameter_length_expansion/param_len_associative_array_total_elements_star
# The ${#map[*]} expansion returns the total number of entries in an associative array.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map=( [a]="1" [b]="2" )
[ "${#map[*]}" -eq 2 ] || fail "associative array star size: want 2, got ${#map[*]}"
echo PASS
exit 0

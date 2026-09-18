#!/usr/bin/env bash
# vybe-test: bash/parameter_length_expansion/param_len_indexed_array_total_elements_star
# The ${#arr[*]} expansion also returns the total number of elements in an indexed array.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=( "alpha" "beta" "gamma" )
[ "${#arr[*]}" -eq 3 ] || fail "array star element count: want 3, got ${#arr[*]}"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/parameter_length_expansion/param_len_indexed_array_total_elements_at
# The ${#arr[@]} expansion returns the total number of elements in an indexed array.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=( "one" "two" "three" "four" )
[ "${#arr[@]}" -eq 4 ] || fail "array element count: want 4, got ${#arr[@]}"

# Sparse array element count
arr[10]="eleventh"
[ "${#arr[@]}" -eq 5 ] || fail "sparse array element count: want 5, got ${#arr[@]}"
echo PASS
exit 0

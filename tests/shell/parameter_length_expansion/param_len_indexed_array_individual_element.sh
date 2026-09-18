#!/usr/bin/env bash
# vybe-test: bash/parameter_length_expansion/param_len_indexed_array_individual_element
# The ${#arr[index]} expansion returns the string character length of element at index.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=( "cat" "elephant" "dog" )
[ "${#arr[0]}" -eq 3 ] || fail "element 0 length: want 3, got ${#arr[0]}"
[ "${#arr[1]}" -eq 8 ] || fail "element 1 length: want 8, got ${#arr[1]}"
[ "${#arr[2]}" -eq 3 ] || fail "element 2 length: want 3, got ${#arr[2]}"
echo PASS
exit 0

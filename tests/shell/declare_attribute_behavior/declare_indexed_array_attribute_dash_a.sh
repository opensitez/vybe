#!/usr/bin/env bash
# vybe-test: bash/declare_attribute_behavior/declare_indexed_array_attribute_dash_a
# The 'declare -a' flag explicitly declares an indexed array.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -a arr=( "elem0" "elem1" )
[ "${#arr[@]}" -eq 2 ] || fail "indexed array length: want 2, got ${#arr[@]}"
arr[5]="elem5"
[ "${#arr[@]}" -eq 3 ] || fail "sparse array length: want 3, got ${#arr[@]}"
[ "${arr[5]}" = "elem5" ] || fail "sparse array element 5 failed: got [${arr[5]}]"
echo PASS
exit 0

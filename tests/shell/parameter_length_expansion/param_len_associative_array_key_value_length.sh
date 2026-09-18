#!/usr/bin/env bash
# vybe-test: bash/parameter_length_expansion/param_len_associative_array_key_value_length
# The ${#map[key]} expansion returns the string character length of the value stored at key.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map=( [greeting]="hello_there" )
[ "${#map[greeting]}" -eq 11 ] || fail "map entry length: want 11, got ${#map[greeting]}"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/indirect_parameter_expansion/indirect_array_keys_expansion_indexed_at
# The ${!arr[@]} expansion expands to the list of set indices in an indexed array.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=( "val0" "val1" "val2" )
indices="${!arr[@]}"
[ "$indices" = "0 1 2" ] || fail "array indices mismatch: want '0 1 2', got [$indices]"
echo PASS
exit 0

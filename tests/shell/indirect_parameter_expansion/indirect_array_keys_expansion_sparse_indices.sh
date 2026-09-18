#!/usr/bin/env bash
# vybe-test: bash/indirect_parameter_expansion/indirect_array_keys_expansion_sparse_indices
# The ${!arr[@]} expansion accurately skips unassigned gaps in sparse indexed arrays.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
sparse_arr[0]="zero"
sparse_arr[5]="five"
sparse_arr[20]="twenty"
indices="${!sparse_arr[@]}"
[ "$indices" = "0 5 20" ] || fail "sparse array indices: want '0 5 20', got [$indices]"
echo PASS
exit 0

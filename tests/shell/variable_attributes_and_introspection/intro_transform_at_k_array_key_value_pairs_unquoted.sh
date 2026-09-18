#!/usr/bin/env bash
# vybe-test: bash/variable_attributes_and_introspection/intro_transform_at_k_array_key_value_pairs_unquoted
# The ${arr[@]@k} parameter transformation expands indexed array elements as alternating index/unquoted-value pairs.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=( [0]="first" [1]="second" )
pairs="${arr[@]@k}"
[ "$pairs" = "0 first 1 second" ] || fail "\${arr[@]@k} output mismatch: got [$pairs]"
echo PASS
exit 0

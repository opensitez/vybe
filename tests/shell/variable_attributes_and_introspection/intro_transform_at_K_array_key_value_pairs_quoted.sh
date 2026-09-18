#!/usr/bin/env bash
# vybe-test: bash/variable_attributes_and_introspection/intro_transform_at_K_array_key_value_pairs_quoted
# The ${arr[@]@K} parameter transformation expands indexed array elements as alternating index/quoted-value pairs.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=( [0]="alpha" [1]="beta" )
pairs="${arr[@]@K}"
case "$pairs" in
    *'0 "alpha" 1 "beta"'*|*'0 "alpha"'*'1 "beta"'*) : ;;
    *) fail "\${arr[@]@K} output mismatch: got [$pairs]" ;;
esac
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/variable_attributes_and_introspection/intro_declare_dash_p_introspects_indexed_array
# Running 'declare -p' on an indexed array displays the full array structure and -a flag.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -a sample_arr=( [0]="item0" [1]="item1" )
output=$(declare -p sample_arr)
case "$output" in
    *"declare -a sample_arr="*'[0]="item0"'*) : ;;
    *) fail "declare -p output missing array details: got [$output]" ;;
esac
echo PASS
exit 0

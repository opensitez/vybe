#!/usr/bin/env bash
# vybe-test: bash/variable_attributes_and_introspection/intro_test_dash_v_evaluates_indexed_array_element_existence
# The [[ -v arr[index] ]] expression introspects whether a specific indexed array element is assigned.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=( [0]="first" [2]="third" )
[[ -v arr[0] ]] || fail "element 0 should exist"
[[ ! -v arr[1] ]] || fail "element 1 should not exist in sparse array"
[[ -v arr[2] ]] || fail "element 2 should exist"
echo PASS
exit 0

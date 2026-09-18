#!/usr/bin/env bash
# vybe-test: bash/parameter_alternate_values/param_alt_indexed_array_element_alternate
# The alternate value expansion applies to individual indexed array elements.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=( "first" )
res0="${arr[0]:+elem0_exists}"
res1="${arr[1]:+elem1_exists}"
[ "$res0" = "elem0_exists" ] || fail "arr[0] alternate failed: got [$res0]"
[ -z "$res1" ] || fail "arr[1] alternate should be empty: got [$res1]"
echo PASS
exit 0

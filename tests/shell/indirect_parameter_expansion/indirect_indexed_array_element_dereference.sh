#!/usr/bin/env bash
# vybe-test: bash/indirect_parameter_expansion/indirect_indexed_array_element_dereference
# The ${!ptr} expansion where ptr="arr[index]" dereferences the indexed array element at index.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=( "apple" "banana" "cherry" )
ptr="arr[1]"
[ "${!ptr}" = "banana" ] || fail "indirect array element access failed: got [${!ptr}]"
ptr="arr[2]"
[ "${!ptr}" = "cherry" ] || fail "indirect array element 2 failed: got [${!ptr}]"
echo PASS
exit 0

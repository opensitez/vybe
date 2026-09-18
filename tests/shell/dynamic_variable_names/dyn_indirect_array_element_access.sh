#!/usr/bin/env bash
# vybe-test: bash/dynamic_variable_names/dyn_indirect_array_element_access
# Indirect expansion ${!pointer} where pointer="arr[index]" retrieves the element at that index.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fruits=( "apple" "banana" "cherry" )
idx=1
pointer="fruits[$idx]"
[ "${!pointer}" = "banana" ] || fail "indirect array element access failed: got [${!pointer}]"
idx=2
pointer="fruits[$idx]"
[ "${!pointer}" = "cherry" ] || fail "indirect array element access at index 2 failed"
echo PASS
exit 0

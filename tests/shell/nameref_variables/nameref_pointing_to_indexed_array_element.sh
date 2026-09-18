#!/usr/bin/env bash
# vybe-test: bash/nameref_variables/nameref_pointing_to_indexed_array_element
# A nameref variable can point directly to a specific element of an indexed array.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=( "elem0" "elem1" "elem2" )
declare -n elem="arr[1]"
[ "$elem" = "elem1" ] || fail "reading array element through nameref failed: got [$elem]"
elem="updated_elem1"
[ "${arr[1]}" = "updated_elem1" ] || fail "array element not mutated through nameref: got [${arr[1]}]"
echo PASS
exit 0

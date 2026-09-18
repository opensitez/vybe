#!/usr/bin/env bash
# vybe-test: bash/nameref_variables/nameref_function_pass_by_reference_indexed_array
# A function using 'local -n' receives an indexed array and appends elements to the caller's array.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
append_element() {
    local -n arr_ref=$1
    local val=$2
    arr_ref+=( "$val" )
}
items=( "alpha" "beta" )
append_element items "gamma"
[ "${#items[@]}" -eq 3 ] || fail "array length after append: want 3, got ${#items[@]}"
[ "${items[2]}" = "gamma" ] || fail "appended element mismatch: got [${items[2]}]"
echo PASS
exit 0

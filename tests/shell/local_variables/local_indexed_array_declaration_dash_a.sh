#!/usr/bin/env bash
# vybe-test: bash/local_variables/local_indexed_array_declaration_dash_a
# Declaring 'local -a arr=( ... )' creates an indexed array scoped exclusively to the function.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr_fn() {
    local -a items=( "one" "two" "three" )
    [ "${#items[@]}" -eq 3 ] || fail "local array length: want 3, got ${#items[@]}"
    [ "${items[1]}" = "two" ] || fail "items[1]: want 'two', got [${items[1]}]"
}
arr_fn
[[ ! -v items ]] || fail "local array items should not exist in outer scope"
echo PASS
exit 0

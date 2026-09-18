#!/usr/bin/env bash
# vybe-test: bash/variable_scope_and_shadowing/scope_shadowing_array_global_with_scalar_local
# A local scalar variable shadows an outer global indexed array during function execution.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=( "item0" "item1" "item2" )
scalar_shadow_fn() {
    local arr="flat_scalar"
    [ "$arr" = "flat_scalar" ] || fail "scalar shadow failed: got [$arr]"
}
scalar_shadow_fn
[ "${#arr[@]}" -eq 3 ] || fail "global array length corrupted: want 3, got ${#arr[@]}"
[ "${arr[1]}" = "item1" ] || fail "global array element corrupted"
echo PASS
exit 0

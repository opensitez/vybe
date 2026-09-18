#!/usr/bin/env bash
# vybe-test: bash/variable_scope_and_shadowing/scope_shadowing_associative_array_with_scalar_local
# A local scalar variable shadows an outer global associative array during function execution.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map=( [key1]="val1" [key2]="val2" )
shadow_assoc_fn() {
    local map="simple_text"
    [ "$map" = "simple_text" ] || fail "scalar shadow of map failed"
}
shadow_assoc_fn
[ "${map[key1]}" = "val1" ] || fail "global map key1 corrupted"
[ "${#map[@]}" -eq 2 ] || fail "global map size corrupted"
echo PASS
exit 0

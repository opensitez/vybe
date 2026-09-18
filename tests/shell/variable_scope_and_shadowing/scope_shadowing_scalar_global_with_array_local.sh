#!/usr/bin/env bash
# vybe-test: bash/variable_scope_and_shadowing/scope_shadowing_scalar_global_with_array_local
# A local indexed array shadows an outer scalar variable during function execution.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
entity="scalar_global"
array_shadow_fn() {
    local -a entity=( "first" "second" )
    [ "${#entity[@]}" -eq 2 ] || fail "local array length: want 2, got ${#entity[@]}"
    [ "${entity[0]}" = "first" ] || fail "entity[0]: want 'first', got [${entity[0]}]"
}
array_shadow_fn
[ "$entity" = "scalar_global" ] || fail "global scalar corrupted: got [$entity]"
echo PASS
exit 0

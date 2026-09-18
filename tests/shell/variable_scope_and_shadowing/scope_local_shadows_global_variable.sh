#!/usr/bin/env bash
# vybe-test: bash/variable_scope_and_shadowing/scope_local_shadows_global_variable
# Declaring a local variable inside a function shadows an outer global variable of the same name.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
target="global_val"
fn() {
    local target="local_val"
    [ "$target" = "local_val" ] || exit 1
}
fn
st=$?
[ "$st" -eq 0 ] || fail "local shadowing failed"
[ "$target" = "global_val" ] || fail "global variable was modified: got [$target]"
echo PASS
exit 0

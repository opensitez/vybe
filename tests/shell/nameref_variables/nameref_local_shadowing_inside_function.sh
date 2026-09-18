#!/usr/bin/env bash
# vybe-test: bash/nameref_variables/nameref_local_shadowing_inside_function
# A function can declare a local nameref with the same name as a global variable, shadowing it cleanly.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
alias_name="global_scalar"
local_target="inner_val"
scope_fn() {
    local -n alias_name=local_target
    alias_name="inner_updated"
}
scope_fn
[ "$alias_name" = "global_scalar" ] || fail "global variable corrupted by local nameref shadow"
[ "$local_target" = "inner_updated" ] || fail "local target not updated through local nameref"
echo PASS
exit 0

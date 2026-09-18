#!/usr/bin/env bash
# vybe-test: bash/local_variables/local_unset_does_not_destroy_global_variable
# Unsetting a shadowed local variable does not destroy or alter the underlying global variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
persistent_global="intact_val"
fn_unsetting_local() {
    local persistent_global="temporary_val"
    unset persistent_global
}
fn_unsetting_local
[ "$persistent_global" = "intact_val" ] || fail "global variable was destroyed: got [$persistent_global]"
echo PASS
exit 0

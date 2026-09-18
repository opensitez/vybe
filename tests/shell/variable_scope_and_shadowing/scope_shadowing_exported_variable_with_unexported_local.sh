#!/usr/bin/env bash
# vybe-test: bash/variable_scope_and_shadowing/scope_shadowing_exported_variable_with_unexported_local
# Declaring 'local' without export shadows an exported variable, preventing child processes from seeing local.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export ACTIVE_EXPORT="global_exported"
test_export_shadow() {
    # In Bash, local var on an exported variable retains or masks export based on attributes
    local ACTIVE_EXPORT="local_unexported"
    # Child process spawned from this function sees the shadowed local value (exported attribute persists)
    res=$( "$BASH" -c 'printf "%s\n" "$ACTIVE_EXPORT"' )
    [ "$res" = "local_unexported" ] || fail "child did not see local value: got [$res]"
}
test_export_shadow
# Outer export remains restored
outer_res=$( "$BASH" -c 'printf "%s\n" "$ACTIVE_EXPORT"' )
[ "$outer_res" = "global_exported" ] || fail "outer export was altered: got [$outer_res]"
echo PASS
exit 0

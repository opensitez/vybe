#!/usr/bin/env bash
# vybe-test: bash/variable_scope_and_shadowing/scope_temporary_command_prefix_shadows_during_command
# A temporary command prefix variable shadows any existing variable for the duration of that command only.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
inspect_scoped_var() {
    [ "$CONFIG_VAR" = "transient_override" ] || exit 1
}
CONFIG_VAR="permanent_state"
CONFIG_VAR="transient_override" inspect_scoped_var
st=$?
[ "$st" -eq 0 ] || fail "transient prefix failed to shadow variable during command execution"
[ "$CONFIG_VAR" = "permanent_state" ] || fail "permanent variable altered by prefix: got [$CONFIG_VAR]"
echo PASS
exit 0

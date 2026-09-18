#!/usr/bin/env bash
# vybe-test: bash/scalar_variable_assignment/scalar_command_prefix_temporary_assignment
# Variable assignment prefixing a command exports the variable temporarily to that command's environment.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
inspect_env() {
    [ "$TEMP_SETTING" = "active" ] || exit 1
}
TEMP_SETTING="active" inspect_env
st=$?
[ "$st" -eq 0 ] || fail "temporary prefix assignment was not visible to invoked command"
echo PASS
exit 0

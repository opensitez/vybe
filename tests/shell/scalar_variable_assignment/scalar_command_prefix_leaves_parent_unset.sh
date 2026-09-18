#!/usr/bin/env bash
# vybe-test: bash/scalar_variable_assignment/scalar_command_prefix_leaves_parent_unset
# Temporary prefix assignment does not persist in the parent shell environment after the command finishes.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset TRANSIENT_VAR
dummy_cmd() { :; }
TRANSIENT_VAR="hello" dummy_cmd
[ -z "$TRANSIENT_VAR" ] || fail "transient variable leaked to parent shell: got [$TRANSIENT_VAR]"
echo PASS
exit 0

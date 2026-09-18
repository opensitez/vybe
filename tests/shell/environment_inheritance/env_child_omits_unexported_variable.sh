#!/usr/bin/env bash
# vybe-test: bash/environment_inheritance/env_child_omits_unexported_variable
# Variables that are not exported are not placed into the process environment table and are omitted in children.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
UNEXPORTED_PARENT_VAR="local_only_state"
child_val=$( "$BASH" -c 'printf "%s\n" "$UNEXPORTED_PARENT_VAR"' )
[ -z "$child_val" ] || fail "child process inherited unexported variable: got [$child_val]"
echo PASS
exit 0

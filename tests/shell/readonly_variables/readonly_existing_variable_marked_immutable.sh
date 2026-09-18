#!/usr/bin/env bash
# vybe-test: bash/readonly_variables/readonly_existing_variable_marked_immutable
# An existing variable can be marked readonly via 'readonly var' without re-specifying its value.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
dynamic_first="constructed_data"
readonly dynamic_first
[ "$dynamic_first" = "constructed_data" ] || fail "readonly altered value: got [$dynamic_first]"
( dynamic_first="overwrite" ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "variable marked readonly should reject further writes"
echo PASS
exit 0

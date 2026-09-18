#!/usr/bin/env bash
# vybe-test: bash/dynamic_variable_names/dyn_unset_via_variable_identifier_reference
# Passing a variable containing a target variable name to unset removes the target variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
victim="active_value"
var_to_delete="victim"
unset "$var_to_delete"
[[ ! -v victim ]] || fail "dynamic unset failed to remove victim variable"
echo PASS
exit 0

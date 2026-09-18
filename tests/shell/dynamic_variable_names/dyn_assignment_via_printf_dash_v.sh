#!/usr/bin/env bash
# vybe-test: bash/dynamic_variable_names/dyn_assignment_via_printf_dash_v
# The 'printf -v "$name"' builtin dynamically assigns formatted output into the named variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
var_name="dynamically_created_var"
printf -v "$var_name" "value_%d" 42
[ "$dynamically_created_var" = "value_42" ] || fail "printf -v dynamic assignment failed"
echo PASS
exit 0

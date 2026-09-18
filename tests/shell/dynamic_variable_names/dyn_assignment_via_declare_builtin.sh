#!/usr/bin/env bash
# vybe-test: bash/dynamic_variable_names/dyn_assignment_via_declare_builtin
# The declare builtin evaluates dynamically constructed "name=value" assignment strings.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
key_name="dynamic_dec_key"
declare "$key_name=computed_value"
[ "$dynamic_dec_key" = "computed_value" ] || fail "declare dynamic assignment failed"
echo PASS
exit 0

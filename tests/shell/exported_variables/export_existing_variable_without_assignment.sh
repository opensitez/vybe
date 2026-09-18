#!/usr/bin/env bash
# vybe-test: bash/exported_variables/export_existing_variable_without_assignment
# An existing variable can be exported with 'export var' without assigning a new value.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
pre_set="pre_existing_data"
export pre_set
child_res=$( "$BASH" -c 'printf "%s\n" "$pre_set"' )
[ "$child_res" = "pre_existing_data" ] || fail "exported existing variable failed: got [$child_res]"
echo PASS
exit 0

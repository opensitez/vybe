#!/usr/bin/env bash
# vybe-test: bash/exported_variables/export_dash_n_removes_export_attribute
# The 'export -n' flag unexports a variable, preventing it from passing to future child processes.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export DEMOTED="was_exported"
export -n DEMOTED
child_saw=$( "$BASH" -c 'printf "%s\n" "$DEMOTED"' )
[ -z "$child_saw" ] || fail "export -n failed to unexport: child saw [$child_saw]"
[ "$DEMOTED" = "was_exported" ] || fail "variable value was lost: got [$DEMOTED]"
echo PASS
exit 0

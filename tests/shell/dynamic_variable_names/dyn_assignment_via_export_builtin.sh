#!/usr/bin/env bash
# vybe-test: bash/dynamic_variable_names/dyn_assignment_via_export_builtin
# The export builtin evaluates dynamically constructed variable assignments and marks them for export.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
exp_name="DYNAMIC_EXPORT_KEY"
export "$exp_name=exported_dynamically"
child_saw=$( "$BASH" -c 'printf "%s\n" "$DYNAMIC_EXPORT_KEY"' )
[ "$child_saw" = "exported_dynamically" ] || fail "export dynamic assignment failed: got [$child_saw]"
echo PASS
exit 0

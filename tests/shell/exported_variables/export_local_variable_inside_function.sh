#!/usr/bin/env bash
# vybe-test: bash/exported_variables/export_local_variable_inside_function
# Exporting a local variable inside a function makes it visible to children invoked from that function.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fn_with_export() {
    local LOCAL_EXP="scoped_export"
    export LOCAL_EXP
    "$BASH" -c 'printf "%s\n" "$LOCAL_EXP"'
}
child_saw=$(fn_with_export)
[ "$child_saw" = "scoped_export" ] || fail "local export not visible to child of function: got [$child_saw]"
echo PASS
exit 0

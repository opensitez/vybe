#!/usr/bin/env bash
# vybe-test: bash/subshell_grouping/subshell_export_variable_isolation
# Exporting a variable inside a subshell does not export it in the parent shell environment.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset EXPORTED_INSIDE
(
    export EXPORTED_INSIDE="child_val"
    [ "$EXPORTED_INSIDE" = "child_val" ] || exit 1
)
[ -z "$EXPORTED_INSIDE" ] || fail "variable leaked into parent environment: got [$EXPORTED_INSIDE]"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/exported_variables/export_child_process_modifications_do_not_leak_back
# Modifications to an exported variable inside a child process do not affect the parent shell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export CONTROL_DATA="parent_origin"
"$BASH" -c 'CONTROL_DATA="child_mutation"'
[ "$CONTROL_DATA" = "parent_origin" ] || fail "child process modified parent exported variable: got [$CONTROL_DATA]"
echo PASS
exit 0

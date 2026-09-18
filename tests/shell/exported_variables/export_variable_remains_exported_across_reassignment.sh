#!/usr/bin/env bash
# vybe-test: bash/exported_variables/export_variable_remains_exported_across_reassignment
# Once exported, simple reassignments update the variable and keep it exported without re-calling export.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export STICKY_EXP="version1"
STICKY_EXP="version2"
res=$( "$BASH" -c 'printf "%s\n" "$STICKY_EXP"' )
[ "$res" = "version2" ] || fail "reassigned variable lost export attribute: got [$res]"
echo PASS
exit 0

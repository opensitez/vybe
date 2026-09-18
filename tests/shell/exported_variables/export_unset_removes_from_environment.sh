#!/usr/bin/env bash
# vybe-test: bash/exported_variables/export_unset_removes_from_environment
# Unsetting an exported variable removes it completely from both shell and child environment.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export DESTROY_ME="temporary_env_data"
unset DESTROY_ME
child_saw=$( "$BASH" -c 'printf "%s\n" "$DESTROY_ME"' )
[ -z "$child_saw" ] || fail "unset failed to remove variable from environment: child saw [$child_saw]"
echo PASS
exit 0

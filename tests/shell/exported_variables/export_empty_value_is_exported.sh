#!/usr/bin/env bash
# vybe-test: bash/exported_variables/export_empty_value_is_exported
# An exported variable whose value is empty is present and set in the child process environment.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export EMPTY_EXP=""
is_set=$( "$BASH" -c '[[ -v EMPTY_EXP ]] && echo "yes" || echo "no"' )
[ "$is_set" = "yes" ] || fail "empty exported variable not set in child environment: got [$is_set]"
echo PASS
exit 0

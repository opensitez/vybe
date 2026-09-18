#!/usr/bin/env bash
# vybe-test: bash/environment_inheritance/env_empty_string_variable_inherited
# An exported environment variable set to empty string is inherited as defined and set in child processes.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export EMPTY_INHERITED=""
res=$( "$BASH" -c '[[ -v EMPTY_INHERITED ]] && printf "%s" "set_and_empty"' )
[ "$res" = "set_and_empty" ] || fail "empty environment variable not inherited as set"
echo PASS
exit 0

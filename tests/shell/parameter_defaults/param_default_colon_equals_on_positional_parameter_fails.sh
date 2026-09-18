#!/usr/bin/env bash
# vybe-test: bash/parameter_defaults/param_default_colon_equals_on_positional_parameter_fails
# Attempting to assign default to a positional parameter via '${1:=default}' produces an error.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set --
( : "${1:=cannot_assign_positional}" ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "\${1:=default} should return non-zero exit code"
echo PASS
exit 0

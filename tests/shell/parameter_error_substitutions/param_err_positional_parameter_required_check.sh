#!/usr/bin/env bash
# vybe-test: bash/parameter_error_substitutions/param_err_positional_parameter_required_check
# The ${1:?error} syntax validates required positional parameters, exiting non-zero when missing.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set --
( : "${1:?positional parameter 1 is mandatory}" ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "missing positional parameter 1 should trigger non-zero exit"

set -- "provided_arg"
val="${1:?should_succeed}"
[ "$val" = "provided_arg" ] || fail "positional parameter expansion failed: got [$val]"
echo PASS
exit 0

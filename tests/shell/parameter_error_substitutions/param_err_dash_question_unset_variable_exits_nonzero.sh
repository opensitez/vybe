#!/usr/bin/env bash
# vybe-test: bash/parameter_error_substitutions/param_err_dash_question_unset_variable_exits_nonzero
# The ${var?error} syntax without colon on an unset variable writes error to stderr and exits non-zero.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset totally_unset
( : "${totally_unset?variable is missing}" ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "unset variable with ? operator should return non-zero exit code"
echo PASS
exit 0

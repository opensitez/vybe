#!/usr/bin/env bash
# vybe-test: bash/parameter_error_substitutions/param_err_colon_question_unset_variable_exits_nonzero
# The ${var:?error} syntax on an unset variable writes error to stderr and exits with a non-zero code.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset missing_var
( : "${missing_var:?missing_var is required}" ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "unset variable should trigger non-zero exit with :? operator"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/parameter_error_substitutions/param_err_colon_question_null_variable_exits_nonzero
# The ${var:?error} syntax on a null (empty string) variable triggers an error and non-zero exit status.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
empty_var=""
( : "${empty_var:?empty_var cannot be null}" ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "null variable should trigger non-zero exit with :? operator"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/parameter_error_substitutions/param_err_dash_question_null_variable_succeeds_empty
# The ${var?error} syntax without colon on a null variable expands to empty without triggering an error.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
null_var=""
val="${null_var?should_not_trigger_on_null}"
st=$?
[ "$st" -eq 0 ] || fail "null variable with ? operator should not fail"
[ -z "$val" ] || fail "null variable should expand to empty string: got [$val]"
echo PASS
exit 0

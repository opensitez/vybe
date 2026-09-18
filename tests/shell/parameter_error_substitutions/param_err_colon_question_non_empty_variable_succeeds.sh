#!/usr/bin/env bash
# vybe-test: bash/parameter_error_substitutions/param_err_colon_question_non_empty_variable_succeeds
# The ${var:?error} syntax on a non-empty variable expands cleanly to var's value without error.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
valid_var="configured_token"
val="${valid_var:?should_not_trigger}"
[ "$val" = "configured_token" ] || fail "expansion mismatch: got [$val]"
echo PASS
exit 0

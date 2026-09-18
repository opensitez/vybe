#!/usr/bin/env bash
# vybe-test: bash/parameter_error_substitutions/param_err_dash_question_non_empty_variable_succeeds
# The ${var?error} syntax without colon on a non-empty variable expands cleanly to var's value.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
filled="active_value"
val="${filled?error_text}"
[ "$val" = "active_value" ] || fail "expansion mismatch: got [$val]"
echo PASS
exit 0

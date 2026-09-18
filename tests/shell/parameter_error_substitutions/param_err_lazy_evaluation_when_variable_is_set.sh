#!/usr/bin/env bash
# vybe-test: bash/parameter_error_substitutions/param_err_lazy_evaluation_when_variable_is_set
# If the variable is set and non-empty, command substitutions in error word are never executed.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
valid_var="present"
executed=0
val="${valid_var:?$(executed=1; echo 'err')}"
[ "$val" = "present" ] || fail "expansion mismatch: got [$val]"
[ "$executed" -eq 0 ] || fail "error word command substitution was eagerly executed"
echo PASS
exit 0

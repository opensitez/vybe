#!/usr/bin/env bash
# vybe-test: bash/parameter_error_substitutions/param_err_nameref_target_error_propagation
# Error substitution on a nameref pointing to an unset target reports an error on that target.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset phantom_target
declare -n ref=phantom_target
( : "${ref:?target variable is missing}" ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "nameref pointing to unset variable should trigger error with :?"
echo PASS
exit 0

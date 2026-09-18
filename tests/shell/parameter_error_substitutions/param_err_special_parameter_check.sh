#!/usr/bin/env bash
# vybe-test: bash/parameter_error_substitutions/param_err_special_parameter_check
# Error substitution can be evaluated on special parameters like $# and $?.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "arg"
val="${#:?count is required}"
[ "$val" -eq 1 ] || fail "count check failed: got $val"

( exit 0 )
status="${?:?exit status required}"
[ "$status" -eq 0 ] || fail "status check failed: got $status"
echo PASS
exit 0

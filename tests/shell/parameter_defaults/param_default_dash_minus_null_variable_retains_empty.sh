#!/usr/bin/env bash
# vybe-test: bash/parameter_defaults/param_default_dash_minus_null_variable_retains_empty
# The ${var-default} syntax without colon expands to the empty string when var is set to null.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
null_var=""
val="${null_var-fallback_value}"
[ -z "$val" ] || fail "null variable should expand to empty string with - operator: got [$val]"
echo PASS
exit 0

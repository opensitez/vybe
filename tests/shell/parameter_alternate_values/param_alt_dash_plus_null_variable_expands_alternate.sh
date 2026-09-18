#!/usr/bin/env bash
# vybe-test: bash/parameter_alternate_values/param_alt_dash_plus_null_variable_expands_alternate
# The ${var+alt} syntax without colon expands to 'alt' even when var is set to null (empty string).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
null_var=""
res="${null_var+is_defined}"
[ "$res" = "is_defined" ] || fail "null variable should expand to alternate with + operator: got [$res]"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/parameter_alternate_values/param_alt_dash_plus_non_empty_variable
# The ${var+alt} syntax without colon expands to 'alt' when var is non-empty.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
var="val"
res="${var+alt_result}"
[ "$res" = "alt_result" ] || fail "expansion mismatch: got [$res]"
echo PASS
exit 0

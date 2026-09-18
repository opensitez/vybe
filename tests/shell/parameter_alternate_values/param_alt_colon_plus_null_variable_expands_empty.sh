#!/usr/bin/env bash
# vybe-test: bash/parameter_alternate_values/param_alt_colon_plus_null_variable_expands_empty
# The ${var:+alt} syntax expands to empty when var is set to null (empty string).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
empty_var=""
res="${empty_var:+alternate_payload}"
[ -z "$res" ] || fail "null variable should expand to empty with :+ operator: got [$res]"
echo PASS
exit 0

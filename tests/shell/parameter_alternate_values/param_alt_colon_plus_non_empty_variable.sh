#!/usr/bin/env bash
# vybe-test: bash/parameter_alternate_values/param_alt_colon_plus_non_empty_variable
# The ${var:+alt} syntax expands to 'alt' when var is set and non-empty.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
var="active_val"
res="${var:+alternate_payload}"
[ "$res" = "alternate_payload" ] || fail "expansion mismatch: got [$res]"
echo PASS
exit 0

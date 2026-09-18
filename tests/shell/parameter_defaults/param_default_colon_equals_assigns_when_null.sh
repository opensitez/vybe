#!/usr/bin/env bash
# vybe-test: bash/parameter_defaults/param_default_colon_equals_assigns_when_null
# The ${var:=default} syntax assigns default to var if var is set to null (empty string).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
empty_slot=""
res="${empty_slot:=replacement_value}"
[ "$res" = "replacement_value" ] || fail "expansion mismatch: got [$res]"
[ "$empty_slot" = "replacement_value" ] || fail "null variable was not assigned by := operator: got [$empty_slot]"
echo PASS
exit 0

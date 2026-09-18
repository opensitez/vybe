#!/usr/bin/env bash
# vybe-test: bash/parameter_defaults/param_default_colon_equals_assigns_when_unset
# The ${var:=default} syntax assigns default to var if var is unset, and expands to default.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset assigned_slot
res="${assigned_slot:=permanent_default}"
[ "$res" = "permanent_default" ] || fail "expansion mismatch: got [$res]"
[ "$assigned_slot" = "permanent_default" ] || fail "variable was not assigned by := operator: got [$assigned_slot]"
echo PASS
exit 0

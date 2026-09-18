#!/usr/bin/env bash
# vybe-test: bash/parameter_alternate_values/param_alt_colon_plus_unset_variable_expands_empty
# The ${var:+alt} syntax expands to empty when var is unset.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset missing_var
res="${missing_var:+alternate_payload}"
[ -z "$res" ] || fail "unset variable should expand to empty with :+ operator: got [$res]"
echo PASS
exit 0

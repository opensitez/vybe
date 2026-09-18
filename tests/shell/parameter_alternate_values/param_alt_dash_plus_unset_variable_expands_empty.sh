#!/usr/bin/env bash
# vybe-test: bash/parameter_alternate_values/param_alt_dash_plus_unset_variable_expands_empty
# The ${var+alt} syntax without colon expands to empty when var is unset.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset completely_unset
res="${completely_unset+is_defined}"
[ -z "$res" ] || fail "unset variable should expand to empty with + operator: got [$res]"
echo PASS
exit 0

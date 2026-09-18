#!/usr/bin/env bash
# vybe-test: bash/parameter_defaults/param_default_dash_equals_preserves_null_without_assignment
# The ${var=default} syntax without colon expands to empty and does NOT assign when var is null.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
null_slot=""
res="${null_slot=should_not_assign}"
[ -z "$res" ] || fail "expansion should be empty for null var with = operator: got [$res]"
[ -z "$null_slot" ] || fail "null variable was modified by = operator: got [$null_slot]"
echo PASS
exit 0

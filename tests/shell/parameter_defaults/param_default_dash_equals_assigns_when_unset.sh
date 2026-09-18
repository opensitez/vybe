#!/usr/bin/env bash
# vybe-test: bash/parameter_defaults/param_default_dash_equals_assigns_when_unset
# The ${var=default} syntax without colon assigns default to var ONLY if var is unset.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset fresh_entry
res="${fresh_entry=assigned_on_unset}"
[ "$res" = "assigned_on_unset" ] || fail "expansion mismatch: got [$res]"
[ "$fresh_entry" = "assigned_on_unset" ] || fail "unset variable was not assigned: got [$fresh_entry]"
echo PASS
exit 0

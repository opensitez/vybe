#!/usr/bin/env bash
# vybe-test: bash/parameter_defaults/param_default_dash_minus_unset_variable
# The ${var-default} syntax without colon expands to default when var is unset.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset unassigned_val
val="${unassigned_val-fallback_for_unset}"
[ "$val" = "fallback_for_unset" ] || fail "unset fallback without colon failed: got [$val]"
echo PASS
exit 0

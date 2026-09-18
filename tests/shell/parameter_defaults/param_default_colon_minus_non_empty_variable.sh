#!/usr/bin/env bash
# vybe-test: bash/parameter_defaults/param_default_colon_minus_non_empty_variable
# The ${var:-default} syntax expands to var itself when var is non-empty.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
active_var="custom_data"
val="${active_var:-default_text}"
[ "$val" = "custom_data" ] || fail "non-empty variable overwritten by default: got [$val]"
echo PASS
exit 0

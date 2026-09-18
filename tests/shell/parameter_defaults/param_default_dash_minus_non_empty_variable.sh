#!/usr/bin/env bash
# vybe-test: bash/parameter_defaults/param_default_dash_minus_non_empty_variable
# The ${var-default} syntax without colon expands to var's value when var is non-empty.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set_var="original_text"
val="${set_var-fallback_value}"
[ "$val" = "original_text" ] || fail "non-empty variable overridden: got [$val]"
echo PASS
exit 0

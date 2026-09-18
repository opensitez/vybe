#!/usr/bin/env bash
# vybe-test: bash/parameter_defaults/param_default_colon_minus_null_variable
# The ${var:-default} syntax expands to default when var is set but null (empty string).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
empty_var=""
val="${empty_var:-default_for_empty}"
[ "$val" = "default_for_empty" ] || fail "null fallback failed: got [$val]"
[ -z "$empty_var" ] || fail "empty_var should remain empty"
echo PASS
exit 0

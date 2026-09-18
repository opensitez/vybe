#!/usr/bin/env bash
# vybe-test: bash/parameter_defaults/param_default_colon_minus_unset_variable
# The ${var:-default} syntax expands to default when var is unset.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset missing_var
val="${missing_var:-default_text}"
[ "$val" = "default_text" ] || fail "unset fallback failed: got [$val]"
[[ ! -v missing_var ]] || fail "colon-minus should not assign variable"
echo PASS
exit 0

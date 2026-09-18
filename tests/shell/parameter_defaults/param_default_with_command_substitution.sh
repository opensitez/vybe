#!/usr/bin/env bash
# vybe-test: bash/parameter_defaults/param_default_with_command_substitution
# The default word expands command substitutions when evaluated for unset or null parameters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset missing_target
res="${missing_target:-$(printf 'computed_fallback\n')}"
[ "$res" = "computed_fallback" ] || fail "command substitution default failed: got [$res]"
echo PASS
exit 0

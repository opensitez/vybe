#!/usr/bin/env bash
# vybe-test: bash/parameter_defaults/param_default_nested_chain_three_deep
# Nested parameter defaults evaluate recursively through a fallback chain.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset opt_a opt_b
opt_c="third_level_value"
res="${opt_a:-${opt_b:-${opt_c:-ultimate_fallback}}}"
[ "$res" = "third_level_value" ] || fail "nested default failed: got [$res]"

unset opt_c
res_ultimate="${opt_a:-${opt_b:-${opt_c:-ultimate_fallback}}}"
[ "$res_ultimate" = "ultimate_fallback" ] || fail "ultimate fallback failed: got [$res_ultimate]"
echo PASS
exit 0

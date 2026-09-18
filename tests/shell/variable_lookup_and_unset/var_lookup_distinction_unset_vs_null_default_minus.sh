#!/usr/bin/env bash
# vybe-test: bash/variable_lookup_and_unset/var_lookup_distinction_unset_vs_null_default_minus
# The ${var-default} operator substitutes default only when var is unset, but retains empty when null.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset completely_unset
null_var=""
[ "${completely_unset-fallback}" = "fallback" ] || fail "unset fallback failed"
[ "${null_var-fallback}" = "" ] || fail "null variable should NOT use fallback with single minus"
echo PASS
exit 0

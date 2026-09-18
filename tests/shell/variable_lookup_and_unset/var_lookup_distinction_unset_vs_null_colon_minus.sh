#!/usr/bin/env bash
# vybe-test: bash/variable_lookup_and_unset/var_lookup_distinction_unset_vs_null_colon_minus
# The ${var:-default} operator substitutes default when var is either unset OR set to null.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset completely_unset
null_var=""
[ "${completely_unset:-fallback}" = "fallback" ] || fail "unset fallback failed"
[ "${null_var:-fallback}" = "fallback" ] || fail "null variable should use fallback with colon minus"
echo PASS
exit 0

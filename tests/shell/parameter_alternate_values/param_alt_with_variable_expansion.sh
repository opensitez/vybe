#!/usr/bin/env bash
# vybe-test: bash/parameter_alternate_values/param_alt_with_variable_expansion
# The alternate word can dynamically expand another variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
trigger="ok"
payload="injected_data"
res="${trigger:+$payload}"
[ "$res" = "injected_data" ] || fail "variable expansion in alternate failed: got [$res]"
echo PASS
exit 0

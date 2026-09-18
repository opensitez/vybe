#!/usr/bin/env bash
# vybe-test: bash/nameref_variables/nameref_creates_target_if_initially_unset
# Assigning through a nameref whose target does not yet exist creates the target variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset future_variable
declare -n ref=future_variable
ref="instantiated_via_ref"
[ "$future_variable" = "instantiated_via_ref" ] || fail "target variable not instantiated: got [$future_variable]"
echo PASS
exit 0

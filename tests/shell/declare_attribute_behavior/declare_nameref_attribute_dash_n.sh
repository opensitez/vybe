#!/usr/bin/env bash
# vybe-test: bash/declare_attribute_behavior/declare_nameref_attribute_dash_n
# The 'declare -n' flag creates a nameref variable that aliases another variable name.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
target_data="initial"
declare -n alias_var=target_data
alias_var="updated_via_alias"
[ "$target_data" = "updated_via_alias" ] || fail "mutation through nameref failed"
[ "$alias_var" = "updated_via_alias" ] || fail "reading nameref failed"
echo PASS
exit 0

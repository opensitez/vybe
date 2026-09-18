#!/usr/bin/env bash
# vybe-test: bash/conditional_expression_syntax/cond_variable_is_nameref_dash_R
# The -R var operator tests whether a variable has been declared with the nameref attribute.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
target_var="underlying"
declare -n ref_var=target_var
regular_var="ordinary"
[[ -R ref_var ]] || fail "ref_var should test positive with -R"
[[ -R regular_var ]] && fail "regular_var should test negative with -R"
echo PASS
exit 0

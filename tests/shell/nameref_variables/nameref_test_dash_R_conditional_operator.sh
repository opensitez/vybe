#!/usr/bin/env bash
# vybe-test: bash/nameref_variables/nameref_test_dash_R_conditional_operator
# The [[ -R var ]] conditional operator evaluates to true if var has the nameref attribute.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
regular_var="ordinary"
declare -n ref=regular_var
[[ -R ref ]] || fail "[[ -R ref ]] should evaluate to true for nameref"
[[ ! -R regular_var ]] || fail "[[ -R regular_var ]] should evaluate to false for standard variable"
echo PASS
exit 0

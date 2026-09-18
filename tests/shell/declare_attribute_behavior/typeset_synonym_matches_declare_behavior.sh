#!/usr/bin/env bash
# vybe-test: bash/declare_attribute_behavior/typeset_synonym_matches_declare_behavior
# The 'typeset' builtin is an exact synonym for 'declare' supporting the same attributes.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
typeset -i math_var="20 + 30"
[ "$math_var" -eq 50 ] || fail "typeset -i failed: got $math_var"
typeset -u upper_var="typeset_case"
[ "$upper_var" = "TYPESET_CASE" ] || fail "typeset -u failed: got [$upper_var]"
echo PASS
exit 0

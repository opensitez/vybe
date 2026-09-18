#!/usr/bin/env bash
# vybe-test: bash/variable_naming_rules/var_name_case_sensitivity_three_variants
# Variable names in Bash are case-sensitive; 'VAR', 'Var', and 'var' are three distinct variables.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
VAR="UPPER"
Var="Camel"
var="lower"
[ "$VAR" = "UPPER" ] || fail "VAR mismatch: got [$VAR]"
[ "$Var" = "Camel" ] || fail "Var mismatch: got [$Var]"
[ "$var" = "lower" ] || fail "var mismatch: got [$var]"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/conditional_expression_syntax/cond_variable_set_dash_v
# The -v var operator tests whether a variable has been set, regardless of whether its value is empty.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
defined_empty=""
defined_val="data"
unset not_defined
[[ -v defined_empty ]] || fail "defined_empty should test positive with -v"
[[ -v defined_val ]] || fail "defined_val should test positive with -v"
[[ -v not_defined ]] && fail "not_defined should test negative with -v"
echo PASS
exit 0

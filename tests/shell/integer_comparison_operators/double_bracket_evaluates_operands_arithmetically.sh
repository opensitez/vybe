#!/usr/bin/env bash
# vybe-test: bash/integer_comparison_operators/double_bracket_evaluates_operands_arithmetically
# Inside [[ ]] the operands of -eq and friends are arithmetic expressions:
# expressions are computed, bare names are looked up, unknown names are 0.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ 1+1 -eq 2 ]] || fail "expressions are evaluated"
x=5
[[ x -eq 5 ]] || fail "bare variable name is looked up"
[[ x*2 -gt 9 ]] || fail "expression with a variable"
unset abc
[[ abc -eq 0 ]] || fail "unset name evaluates to 0"
echo PASS
exit 0

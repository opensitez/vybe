#!/usr/bin/env bash
# vybe-test: bash/conditional_expression_syntax/cond_short_circuit_logical_or
# The '||' operator in [[ ... ]] short-circuits: if the left operand is true, the right operand is not evaluated.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
executed_right=0
[[ 1 -eq 1 || $(( executed_right++ )) -eq 0 ]]
[ "$executed_right" -eq 0 ] || fail "right operand evaluated despite true left in ||: got $executed_right"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/short_circuit_evaluation/test_a_and_o_evaluate_every_operand
# [ ] is a command: all its arguments are expanded before it runs, so -a and
# -o cannot short-circuit, unlike && inside [[ ]].
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
[ 1 -eq 2 -a $((x++)) -eq 0 ]
[ "$x" -eq 1 ] || fail "[ -a ] operand was not expanded, x=$x"
y=0
[[ 1 -eq 2 && $((y++)) -eq 0 ]]
[ "$y" -eq 0 ] || fail "[[ && ]] must skip the right operand, y=$y"
echo PASS
exit 0

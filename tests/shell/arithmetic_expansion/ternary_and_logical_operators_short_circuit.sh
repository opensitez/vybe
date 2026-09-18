#!/usr/bin/env bash
# vybe-test: bash/arithmetic_expansion/ternary_and_logical_operators_short_circuit
# The unevaluated branch is skipped entirely: no error and no side effect.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
[ $((1 ? 2 : 1/0)) -eq 2 ] || fail "ternary true branch"
[ $((0 ? 1/0 : 3)) -eq 3 ] || fail "ternary false branch"
[ $((0 && (x=9))) -eq 0 ] || fail "&& result"
[ $((1 || (x=9))) -eq 1 ] || fail "|| result"
[ "$x" -eq 0 ] || fail "skipped branch must not assign, x=$x"
echo PASS
exit 0

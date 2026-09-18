#!/usr/bin/env bash
# vybe-test: bash/indirect_parameter_expansion/indirect_dereference_in_arithmetic_context
# The indirect parameter expansion ${!ptr} can be evaluated directly in arithmetic expressions.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
score=50
ptr="score"
(( doubled = ${!ptr} * 2 ))
[ "$doubled" -eq 100 ] || fail "arithmetic evaluation with indirect expansion failed: got $doubled"
sum=$(( ${!ptr} + 25 ))
[ "$sum" -eq 75 ] || fail "inline arithmetic calculation failed: got $sum"
echo PASS
exit 0

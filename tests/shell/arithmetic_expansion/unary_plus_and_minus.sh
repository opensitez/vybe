#!/usr/bin/env bash
# vybe-test: bash/arithmetic_expansion/unary_plus_and_minus
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=3
out="$((-x)) $((+x)) $((- -3)) $((-(x+1))) $((- x * 2))"
[ "$out" = "-3 3 3 -4 -6" ] || fail "got [$out]"
echo PASS
exit 0

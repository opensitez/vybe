#!/usr/bin/env bash
# vybe-test: bash/integer_arithmetic/floating_point_literals_are_errors
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$( eval 'echo $((1.5 + 1))' 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "want status 1 got $st"
[[ $msg == *"arithmetic syntax error"* ]] || fail "got [$msg]"
[ $((7/2*2)) -eq 6 ] || fail "integer division discards the fraction"
echo PASS
exit 0

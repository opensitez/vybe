#!/usr/bin/env bash
# vybe-test: bash/integer_arithmetic/no_exponent_compound_assignment_operator
# ** exists but **= does not; the parser sees ** followed by =2.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=3
msg=$( eval 'echo $(( x **= 2 ))' 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "want status 1 got $st"
[[ $msg == *"operand expected"* ]] || fail "got [$msg]"
[ "$x" -eq 3 ] || fail "x must be untouched, got $x"
echo PASS
exit 0

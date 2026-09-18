#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/increments_do_not_cast_strings_beyond_arithmetic
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x='7abc'
[ $((x++)) -eq 7 ] || fail "string-prefixed number treated as 7"
[ "$x" -eq 8 ] || fail "x should become 8"
y='abc'
msg=$( eval 'echo $((y++) )' 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "non-numeric string should fail in post-increment, got $st"
[[ $msg == *"operand expected"* || $msg == *"syntax error"* ]] || fail "message: [$msg]"
echo PASS
exit 0

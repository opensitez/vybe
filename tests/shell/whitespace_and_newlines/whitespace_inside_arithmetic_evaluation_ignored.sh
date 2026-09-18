#!/usr/bin/env bash
# vybe-test: bash/whitespace_and_newlines/whitespace_inside_arithmetic_evaluation_ignored
# Free-form spaces and tabs inside (( ... )) arithmetic evaluation are completely ignored.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
((   x   =   15   +   25   ))
[ "$x" -eq 40 ] || fail "whitespace in arithmetic: want 40, got $x"
echo PASS
exit 0

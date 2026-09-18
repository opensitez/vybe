#!/usr/bin/env bash
# vybe-test: bash/here_string_syntax/here_string_expands_arithmetic_evaluation
# Arithmetic expansions $(( ... )) are computed before the here-string is passed to standard input.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=20
y=30
read -r result <<< "sum: $(( x + y ))"
[ "$result" = "sum: 50" ] || fail "arithmetic expansion in here-string: got [$result]"
echo PASS
exit 0

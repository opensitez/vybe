#!/usr/bin/env bash
# vybe-test: bash/whitespace_and_newlines/newline_acts_as_command_separator
# An unquoted newline separates sequential commands identically to a semicolon.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=10
y=20
sum=$((x + y))
[ "$sum" -eq 30 ] || fail "sum after newline-separated assignments: want 30, got $sum"
echo PASS
exit 0

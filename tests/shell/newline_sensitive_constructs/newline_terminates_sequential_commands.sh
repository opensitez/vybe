#!/usr/bin/env bash
# vybe-test: bash/newline_sensitive_constructs/newline_terminates_sequential_commands
# An unquoted newline terminates a simple command and initiates parsing of the next command.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=1
b=2
sum=$(( a + b ))
[ "$sum" -eq 3 ] || fail "newline terminated commands: want 3, got $sum"
echo PASS
exit 0

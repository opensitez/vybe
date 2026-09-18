#!/usr/bin/env bash
# vybe-test: bash/whitespace_and_newlines/newline_after_logical_and_continues_command
# A newline immediately following '&&' continues the conditional list to the next line.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
true &&
    x=1
[ "$x" -eq 1 ] || fail "newline after &&: want x=1, got $x"
echo PASS
exit 0

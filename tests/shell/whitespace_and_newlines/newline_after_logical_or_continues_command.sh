#!/usr/bin/env bash
# vybe-test: bash/whitespace_and_newlines/newline_after_logical_or_continues_command
# A newline immediately following '||' continues the conditional list to the next line.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
y=0
false ||
    y=99
[ "$y" -eq 99 ] || fail "newline after ||: want y=99, got $y"
echo PASS
exit 0

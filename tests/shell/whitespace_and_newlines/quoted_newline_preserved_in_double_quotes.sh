#!/usr/bin/env bash
# vybe-test: bash/whitespace_and_newlines/quoted_newline_preserved_in_double_quotes
# A newline inside double quotes is treated as a literal character, not a command terminator.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
multiline="line1
line2"
[ "$multiline" = $'line1\nline2' ] || fail "double quoted newline: got [$multiline]"
echo PASS
exit 0

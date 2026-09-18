#!/usr/bin/env bash
# vybe-test: bash/whitespace_and_newlines/quoted_newline_preserved_in_single_quotes
# A newline inside single quotes is treated as a literal character, not a command terminator.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
multiline='first
second'
[ "$multiline" = $'first\nsecond' ] || fail "single quoted newline: got [$multiline]"
echo PASS
exit 0

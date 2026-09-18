#!/usr/bin/env bash
# vybe-test: bash/newline_sensitive_constructs/newline_literal_inside_single_quotes
# An unescaped newline enclosed within single quotes is preserved as a literal newline character.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
s='line1
line2'
[ "$s" = $'line1\nline2' ] || fail "single quote literal newline failed: got [$s]"
echo PASS
exit 0

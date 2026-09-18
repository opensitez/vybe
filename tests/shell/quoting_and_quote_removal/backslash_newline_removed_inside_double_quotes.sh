#!/usr/bin/env bash
# vybe-test: bash/quoting_and_quote_removal/backslash_newline_removed_inside_double_quotes
# In "…" a backslash-newline pair is a line continuation; in '…' it is kept.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
d="a\
b"
[ "$d" = ab ] || fail "double quotes: want [ab] got [$d]"
s='a\
b'
[ "${#s}" -eq 4 ] || fail "single quotes keep backslash and newline: want length 4 got ${#s}"
echo PASS
exit 0

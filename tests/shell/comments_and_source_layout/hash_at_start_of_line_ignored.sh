#!/usr/bin/env bash
# vybe-test: bash/comments_and_source_layout/hash_at_start_of_line_ignored
# A line starting with '#' is a comment and is ignored by the shell parser.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
# This is a full-line comment
x=100
# Another comment in between
[ "$x" -eq 100 ] || fail "x value: want 100, got $x"
echo PASS
exit 0

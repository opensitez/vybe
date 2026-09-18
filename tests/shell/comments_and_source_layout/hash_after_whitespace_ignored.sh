#!/usr/bin/env bash
# vybe-test: bash/comments_and_source_layout/hash_after_whitespace_ignored
# A '#' character preceded by spaces or tabs introduces an inline comment to end of line.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=42    # inline comment after spaces
y=84	# inline comment after tab
[ "$x" -eq 42 ] || fail "x: want 42, got $x"
[ "$y" -eq 84 ] || fail "y: want 84, got $y"
echo PASS
exit 0

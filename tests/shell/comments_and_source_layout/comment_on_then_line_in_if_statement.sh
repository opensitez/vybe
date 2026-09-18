#!/usr/bin/env bash
# vybe-test: bash/comments_and_source_layout/comment_on_then_line_in_if_statement
# A comment immediately following 'then' on the same line is parsed cleanly.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
if true; then # then branch comment
    x=1 # inside block
else # else branch comment
    x=2
fi
[ "$x" -eq 1 ] || fail "if then comment failed: want 1, got $x"
echo PASS
exit 0

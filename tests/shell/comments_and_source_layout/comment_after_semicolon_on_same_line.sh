#!/usr/bin/env bash
# vybe-test: bash/comments_and_source_layout/comment_after_semicolon_on_same_line
# A comment can immediately follow a semicolon command separator with optional spaces.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=10; # set initial x
y=20;# set initial y without space
[ "$x" -eq 10 ] && [ "$y" -eq 20 ] || fail "comment after semicolon failed"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/comments_and_source_layout/comment_after_logical_and_continues_command
# A comment immediately following '&&' on the same line allows continuing the command on next line.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
true && # proceed if previous command succeeded
    x=1
[ "$x" -eq 1 ] || fail "comment after &&: want x=1, got $x"
echo PASS
exit 0

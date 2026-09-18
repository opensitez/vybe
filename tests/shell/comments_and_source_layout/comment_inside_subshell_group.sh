#!/usr/bin/env bash
# vybe-test: bash/comments_and_source_layout/comment_inside_subshell_group
# Comments inside a subshell compound command (...) are ignored.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=1
(
    # inside subshell
    x=2 # modify local copy
    exit 0
)
[ "$x" -eq 1 ] || fail "parent x: want 1, got $x"
echo PASS
exit 0

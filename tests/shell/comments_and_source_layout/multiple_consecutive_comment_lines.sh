#!/usr/bin/env bash
# vybe-test: bash/comments_and_source_layout/multiple_consecutive_comment_lines
# Multiple back-to-back comment lines are ignored without interfering with subsequent code execution.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
# Comment line 1
# Comment line 2
# Comment line 3
# Comment line 4
ans=999
[ "$ans" -eq 999 ] || fail "ans: want 999, got $ans"
echo PASS
exit 0

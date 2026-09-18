#!/usr/bin/env bash
# vybe-test: bash/comments_and_source_layout/comment_after_ampersand_on_same_line
# A comment can immediately follow an ampersand asynchronous operator.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(exit 0)& # run subshell in background
wait $!
st=$?
[ "$st" -eq 0 ] || fail "wait status: want 0, got $st"
echo PASS
exit 0

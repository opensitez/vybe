#!/usr/bin/env bash
# vybe-test: bash/comments_and_source_layout/comment_on_do_line_in_for_loop
# A comment immediately following 'do' on the same line in a loop is parsed cleanly.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
acc=0
for n in 1 2 3; do # begin loop iteration
    acc=$((acc + n)) # accumulate
done # end loop
[ "$acc" -eq 6 ] || fail "loop acc with comments: want 6, got $acc"
echo PASS
exit 0

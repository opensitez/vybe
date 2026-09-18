#!/usr/bin/env bash
# vybe-test: bash/reserved_words_and_keywords/until_keyword_loops_while_false
# The until reserved word repeatedly executes the loop body while the condition command returns false (non-zero).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count=0
until [ "$count" -ge 3 ]; do
    count=$((count + 1))
done
[ "$count" -eq 3 ] || fail "until loop count: want 3, got $count"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/comments_and_source_layout/comment_inside_array_literal_lines
# Comments inside an array literal definition are stripped cleanly between element lines.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
items=(
    apple   # first item
    banana  # second item
    cherry  # third item
)
[ "${#items[@]}" -eq 3 ] || fail "array count: want 3, got ${#items[@]}"
[ "${items[0]}" = "apple" ] || fail "item 0: want 'apple', got [${items[0]}]"
[ "${items[1]}" = "banana" ] || fail "item 1: want 'banana', got [${items[1]}]"
[ "${items[2]}" = "cherry" ] || fail "item 2: want 'cherry', got [${items[2]}]"
echo PASS
exit 0

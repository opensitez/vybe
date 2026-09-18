#!/usr/bin/env bash
# vybe-test: bash/comments_and_source_layout/hash_embedded_in_word_is_literal
# A '#' embedded inside a word without preceding whitespace is a literal character, not a comment.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
word=foo#bar
[ "$word" = "foo#bar" ] || fail "embedded hash: want 'foo#bar', got [$word]"
echo PASS
exit 0

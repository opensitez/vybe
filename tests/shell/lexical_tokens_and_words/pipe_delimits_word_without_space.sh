#!/usr/bin/env bash
# vybe-test: bash/lexical_tokens_and_words/pipe_delimits_word_without_space
# A pipe character immediately adjacent to words acts as pipeline token without spaces.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(printf 'apple'|cat)
[ "$out" = "apple" ] || fail "pipe without space: want 'apple', got [$out]"
echo PASS
exit 0

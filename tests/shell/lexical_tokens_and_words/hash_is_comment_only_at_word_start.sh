#!/usr/bin/env bash
# vybe-test: bash/lexical_tokens_and_words/hash_is_comment_only_at_word_start
# A # inside a word is literal; a # that starts a word begins a comment that
# runs to the end of the line (even inside $(…), which is why the closing
# parenthesis below sits on its own line).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(echo a#b)
[ "$out" = "a#b" ] || fail "embedded hash: want [a#b] got [$out]"
out=$(echo a #b c d
)
[ "$out" = "a" ] || fail "hash after blank starts a comment: want [a] got [$out]"
echo PASS
exit 0

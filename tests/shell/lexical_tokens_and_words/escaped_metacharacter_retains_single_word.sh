#!/usr/bin/env bash
# vybe-test: bash/lexical_tokens_and_words/escaped_metacharacter_retains_single_word
# Escaping metacharacters like semicolon, pipe, and ampersand keeps them in the same word.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- a\;b c\|d e\&f
[ "$#" -eq 3 ] || fail "arg count: want 3, got $#"
[ "$1" = "a;b" ] || fail "arg 1: want 'a;b', got [$1]"
[ "$2" = "c|d" ] || fail "arg 2: want 'c|d', got [$2]"
[ "$3" = "e&f" ] || fail "arg 3: want 'e&f', got [$3]"
echo PASS
exit 0

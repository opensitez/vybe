#!/usr/bin/env bash
# vybe-test: bash/lexical_tokens_and_words/semicolon_delimits_word_without_space
# A semicolon immediately following a word terminates the word and acts as a command separator.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=init
x=foo; y=bar
[ "$x" = "foo" ] || fail "x: want 'foo', got [$x]"
[ "$y" = "bar" ] || fail "y: want 'bar', got [$y]"
echo PASS
exit 0

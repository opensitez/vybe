#!/usr/bin/env bash
# vybe-test: bash/lexical_tokens_and_words/arithmetic_vs_nested_subshell_tokens
# $(( is arithmetic; $( ( is a command substitution containing a subshell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=$((1+2))
[ "$a" = 3 ] || fail "arithmetic: want 3 got [$a]"
b=$( (echo 4) )
[ "$b" = 4 ] || fail "nested subshell: want 4 got [$b]"
echo PASS
exit 0

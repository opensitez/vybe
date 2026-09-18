#!/usr/bin/env bash
# vybe-test: bash/lexical_tokens_and_words/bang_is_ordinary_word_in_arguments
# History expansion is off in scripts, and ! is reserved only at the start of
# a pipeline.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(echo hello!)
[ "$out" = "hello!" ] || fail "want [hello!] got [$out]"
out=$(echo "!" "!!")
[ "$out" = "! !!" ] || fail "want [! !!] got [$out]"
echo PASS
exit 0

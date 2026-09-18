#!/usr/bin/env bash
# vybe-test: bash/lexical_tokens_and_words/metacharacters_end_words_without_spaces
# ; & | ( ) < > terminate a word even with no surrounding blanks.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(echo a;echo b)
[ "$out" = $'a\nb' ] || fail "semicolon: got [$out]"
out=$(echo c&&echo d)
[ "$out" = $'c\nd' ] || fail "and-and: got [$out]"
out=$(echo e|{ read -r v; echo "got:$v"; })
[ "$out" = "got:e" ] || fail "pipe: got [$out]"
echo PASS
exit 0

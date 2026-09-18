#!/usr/bin/env bash
# vybe-test: bash/lexical_tokens_and_words/mixed_quoting_forms_one_word
# Unquoted, ${}, '…' and "…" segments without blanks between them are one word.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
b=X
n=$(count a${b}c'd'"e")
[ "$n" = 1 ] || fail "want 1 argument, got $n"
out=$(echo a${b}c'd'"e")
[ "$out" = "aXcde" ] || fail "want [aXcde] got [$out]"
echo PASS
exit 0

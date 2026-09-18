#!/usr/bin/env bash
# vybe-test: bash/lexical_tokens_and_words/quoted_keyword_is_argument
# A reserved word loses its special meaning when any part of it is quoted:
# "if" in command position is looked up as a command, not parsed as a keyword.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
"if" true 2>/dev/null
st=$?
[ "$st" -eq 127 ] || fail "quoted if should be command-not-found (127), got status $st"
out=$(echo if then fi)
[ "$out" = "if then fi" ] || fail "keywords as arguments: want [if then fi] got [$out]"
echo PASS
exit 0

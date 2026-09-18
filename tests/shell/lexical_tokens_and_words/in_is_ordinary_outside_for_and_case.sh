#!/usr/bin/env bash
# vybe-test: bash/lexical_tokens_and_words/in_is_ordinary_outside_for_and_case
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
in=5
[ "$in" = 5 ] || fail "in as a variable name: got [$in]"
out=$(echo in do done)
[ "$out" = "in do done" ] || fail "want [in do done] got [$out]"
echo PASS
exit 0

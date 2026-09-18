#!/usr/bin/env bash
# vybe-test: bash/lexical_tokens_and_words/keyword_recognized_after_list_operator
# The word after && || ; | is in command position, so keywords are live there.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(true && if true; then echo a; fi; false || while false; do :; done; echo b)
[ "$out" = $'a\nb' ] || fail "want [a\\nb] got [$out]"
echo PASS
exit 0

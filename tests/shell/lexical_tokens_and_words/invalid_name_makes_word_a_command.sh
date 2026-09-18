#!/usr/bin/env bash
# vybe-test: bash/lexical_tokens_and_words/invalid_name_makes_word_a_command
# A word with = whose left side is not a valid name is not an assignment; it
# is executed as a command.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
1x=3 2>/dev/null
st=$?
[ "$st" -eq 127 ] || fail "1x=3 should be command not found (127), got $st"
a-b=3 2>/dev/null
st=$?
[ "$st" -eq 127 ] || fail "a-b=3 should be command not found (127), got $st"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/reserved_words_and_keywords/quoted_reserved_word_is_ordinary_command_name
# A quoted keyword is treated as an ordinary command word and does not start a compound command.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
"if" 2>/dev/null
st=$?
[ "$st" -eq 127 ] || fail "quoted 'if' should be command not found (127), got $st"

"while" 2>/dev/null
st=$?
[ "$st" -eq 127 ] || fail "quoted 'while' should be command not found (127), got $st"
echo PASS
exit 0

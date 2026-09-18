#!/usr/bin/env bash
# vybe-test: bash/lexical_tokens_and_words/double_bracket_requires_whitespace_delimiters
# [[ and ]] are reserved words, not metacharacters, so [[x=x]] is parsed as a single word command.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval '[[1=1]]' 2>/dev/null
st=$?
[ "$st" -eq 127 ] || fail "[[1=1]] without spaces: want 127 (command not found), got $st"
[[ 1 -eq 1 ]]
st=$?
[ "$st" -eq 0 ] || fail "[[ 1 -eq 1 ]] with spaces: want 0, got $st"
echo PASS
exit 0

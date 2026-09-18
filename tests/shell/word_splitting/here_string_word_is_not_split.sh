#!/usr/bin/env bash
# vybe-test: bash/word_splitting/here_string_word_is_not_split
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x='a   b'
IFS= read -r v <<< $x
[ "$v" = 'a   b' ] || fail "got [$v]"
echo PASS
exit 0

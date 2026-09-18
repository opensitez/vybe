#!/usr/bin/env bash
# vybe-test: bash/word_splitting/whitespace_only_value_produces_no_words
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
x='   '
[ "$(count $x)" = 0 ] || fail "unquoted: want 0 got $(count $x)"
[ "$(count "$x")" = 1 ] || fail "quoted: want 1"
echo PASS
exit 0

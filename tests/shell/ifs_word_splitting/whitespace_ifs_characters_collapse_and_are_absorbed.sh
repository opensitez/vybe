#!/usr/bin/env bash
# vybe-test: bash/ifs_word_splitting/whitespace_ifs_characters_collapse_and_are_absorbed
# IFS whitespace around a non-whitespace delimiter belongs to that delimiter.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
IFS=': '
x='a : b'
[ "$(count $x)" = 2 ] || fail "'a : b' want 2 got $(count $x)"
x='a::b'
[ "$(count $x)" = 3 ] || fail "'a::b' want 3 got $(count $x)"
x='a  b'
[ "$(count $x)" = 2 ] || fail "'a  b' want 2 got $(count $x)"
echo PASS
exit 0

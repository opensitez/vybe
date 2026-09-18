#!/usr/bin/env bash
# vybe-test: bash/ifs_word_splitting/newline_only_ifs_splits_lines
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
IFS=$'\n'
x=$'a b\nc d\n'
[ "$(count $x)" = 2 ] || fail "want 2 got $(count $x)"
first() { echo "$1"; }
[ "$(first $x)" = 'a b' ] || fail "spaces must survive: got [$(first $x)]"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/word_splitting/whitespace_ifs_trims_and_collapses
# With the default IFS, leading/trailing whitespace is dropped and any run of
# whitespace is a single separator.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
x=$'  a \t b\n\nc  '
[ "$(count $x)" = 3 ] || fail "want 3 got $(count $x)"
first() { echo "[$1]"; }
[ "$(first $x)" = "[a]" ] || fail "leading blanks must not make an empty word"
echo PASS
exit 0

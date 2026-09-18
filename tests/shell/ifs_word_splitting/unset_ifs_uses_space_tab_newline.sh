#!/usr/bin/env bash
# vybe-test: bash/ifs_word_splitting/unset_ifs_uses_space_tab_newline
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
unset IFS
x=$'a\tb\nc d:e'
[ "$(count $x)" = 4 ] || fail "want 4 got $(count $x)"
echo PASS
exit 0

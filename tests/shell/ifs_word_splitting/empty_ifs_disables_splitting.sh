#!/usr/bin/env bash
# vybe-test: bash/ifs_word_splitting/empty_ifs_disables_splitting
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
IFS=
x='a b:c'
[ "$(count $x)" = 1 ] || fail "want 1 got $(count $x)"
echo PASS
exit 0

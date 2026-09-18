#!/usr/bin/env bash
# vybe-test: bash/ifs_word_splitting/ifs_splits_only_expansion_results_not_literal_text
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
IFS=:
x=a:b
[ "$(count a:b $x)" = 3 ] || fail "literal a:b is one word, \$x is two: want 3 got $(count a:b $x)"
echo PASS
exit 0

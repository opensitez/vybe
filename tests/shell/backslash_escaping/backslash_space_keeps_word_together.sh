#!/usr/bin/env bash
# vybe-test: bash/backslash_escaping/backslash_space_keeps_word_together
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
first() { echo "$1"; }
n=$(count a\ b c)
[ "$n" = 2 ] || fail "want 2 args got $n"
v=$(first a\ b c)
[ "$v" = "a b" ] || fail "want [a b] got [$v]"
echo PASS
exit 0

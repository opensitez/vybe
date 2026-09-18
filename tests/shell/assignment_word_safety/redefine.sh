#!/usr/bin/env bash
# vybe-test: bash/assignment_word_safety/redefine
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=9
b=$((a + 1))
c='x y'
d=("$c")
[ "$a" -eq 9 ] || fail "arithmetic assignment broken"
[ "$b" -eq $((9 + 1)) ] || fail "nested arithmetic broken"
[ "${d[0]}" = "x y" ] || fail "word splitting in assignment not expected"
echo PASS
exit 0

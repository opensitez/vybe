#!/usr/bin/env bash
# vybe-test: bash/assignment_word_safety/stress
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=18
b=$((a + 1))
c='x y'
d=("$c")
[ "$a" -eq 18 ] || fail "arithmetic assignment broken"
[ "$b" -eq $((18 + 1)) ] || fail "nested arithmetic broken"
[ "${d[0]}" = "x y" ] || fail "word splitting in assignment not expected"
echo PASS
exit 0

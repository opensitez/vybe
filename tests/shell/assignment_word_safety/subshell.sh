#!/usr/bin/env bash
# vybe-test: bash/assignment_word_safety/subshell
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=3
b=$((a + 1))
c='x y'
d=("$c")
[ "$a" -eq 3 ] || fail "arithmetic assignment broken"
[ "$b" -eq $((3 + 1)) ] || fail "nested arithmetic broken"
[ "${d[0]}" = "x y" ] || fail "word splitting in assignment not expected"
echo PASS
exit 0

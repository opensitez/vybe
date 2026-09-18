#!/usr/bin/env bash
# vybe-test: bash/ansi_c_quoting/nul_escape_terminates_the_string
# Shell strings cannot hold NUL: \0 ends the value.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=$'a\0b'
[ "$x" = a ] || fail "want [a] got [$x]"
y=$'\0'
[ -z "$y" ] || fail "lone \\0 must be empty, got [$y]"
z=$'\x00tail'
[ -z "$z" ] || fail "\\x00 also terminates, got [$z]"
echo PASS
exit 0

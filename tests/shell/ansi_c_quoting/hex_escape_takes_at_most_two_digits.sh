#!/usr/bin/env bash
# vybe-test: bash/ansi_c_quoting/hex_escape_takes_at_most_two_digits
# \xHH reads one or two hex digits; a third digit is ordinary text. Octal
# \nnn reads up to three digits.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=$'\x411'
[ "$x" = 'A1' ] || fail "\\x411 want [A1] got [$x]"
y=$'\1011'
[ "$y" = 'A1' ] || fail "\\1011 want [A1] got [$y]"
z=$'\x4'; four=$'\x04'
[ "$z" = "$four" ] || fail "single hex digit"
echo PASS
exit 0

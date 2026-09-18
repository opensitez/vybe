#!/usr/bin/env bash
# vybe-test: bash/command_substitution/only_trailing_newlines_are_stripped
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=$(printf 'a\n\nb\n\n\n')
[ "$x" = $'a\n\nb' ] || fail "interior newlines kept, trailing removed: got [$x]"
y=$(printf '  a  ')
[ "$y" = '  a  ' ] || fail "spaces are not stripped: got [$y]"
z=$(printf 'a\n \n')
[ "$z" = $'a\n ' ] || fail "a space after the last newline stops stripping: got [$z]"
echo PASS
exit 0

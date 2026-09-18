#!/usr/bin/env bash
# vybe-test: bash/backslash_escaping/backslash_newline_continues_command
# Unquoted backslash-newline is deleted; the words on both lines belong to
# the same command and remain separate words.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
n=$(count a \
b)
[ "$n" = 2 ] || fail "want 2 args got $n"
out=$(echo a \
b)
[ "$out" = "a b" ] || fail "want [a b] got [$out]"
echo PASS
exit 0

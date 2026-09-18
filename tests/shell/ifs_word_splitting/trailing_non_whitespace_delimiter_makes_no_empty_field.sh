#!/usr/bin/env bash
# vybe-test: bash/ifs_word_splitting/trailing_non_whitespace_delimiter_makes_no_empty_field
# A single trailing delimiter terminates the last field; only a second one
# adds an empty field.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
IFS=:
x=a:
[ "$(count $x)" = 1 ] || fail "a: want 1 got $(count $x)"
x=a::
[ "$(count $x)" = 2 ] || fail "a:: want 2 got $(count $x)"
echo PASS
exit 0

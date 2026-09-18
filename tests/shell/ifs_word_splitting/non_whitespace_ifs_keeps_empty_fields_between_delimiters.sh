#!/usr/bin/env bash
# vybe-test: bash/ifs_word_splitting/non_whitespace_ifs_keeps_empty_fields_between_delimiters
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
IFS=:
x=a::b
[ "$(count $x)" = 3 ] || fail "a::b want 3 got $(count $x)"
x=:a
[ "$(count $x)" = 2 ] || fail "leading delimiter makes an empty first field, want 2 got $(count $x)"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/ifs_word_splitting/temporary_ifs_for_a_single_read
# IFS=: read … applies the new IFS only to that read.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
IFS=: read -r a b <<< 'x:y z'
[ "$a" = x ] && [ "$b" = 'y z' ] || fail "read fields: a=[$a] b=[$b]"
x='p:q r'
[ "$(count $x)" = 2 ] || fail "IFS must be restored afterwards, got $(count $x) words"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/reserved_words_and_keywords/double_bracket_logical_operators_are_keywords
# Within [[ ... ]], && and || are treated as conditional expression logical operators.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ 1 -eq 1 && "a" == "a" ]]
st1=$?
[ "$st1" -eq 0 ] || fail "[[ ... && ... ]]: want 0, got $st1"

[[ 1 -eq 2 || "b" == "b" ]]
st2=$?
[ "$st2" -eq 0 ] || fail "[[ ... || ... ]]: want 0, got $st2"

[[ 1 -eq 2 && "b" == "b" ]]
st3=$?
[ "$st3" -ne 0 ] || fail "[[ false && true ]]: want non-zero, got $st3"
echo PASS
exit 0

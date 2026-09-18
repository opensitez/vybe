#!/usr/bin/env bash
# vybe-test: bash/pattern_comparisons/escaped_glob_characters_match_literally
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ 'a*' == a\* ]] || fail "escaped star matches a literal star"
[[ ab == a\* ]] && fail "escaped star must not act as a wildcard"
[[ 'a?' == a\? ]] || fail "escaped question mark"
[[ '[x]' == \[x\] ]] || fail "escaped brackets"
echo PASS
exit 0

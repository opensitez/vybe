#!/usr/bin/env bash
# vybe-test: bash/empty_strings_and_null_words/case_matches_empty_word_with_empty_or_star_pattern
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=
case $x in "") r=empty ;; *) r=other ;; esac
[ "$r" = empty ] || fail "unquoted empty word must not be an error and must match \"\", got [$r]"
case "$x" in ?) r=one ;; *) r=star ;; esac
[ "$r" = star ] || fail "? needs one char; * matches empty, got [$r]"
echo PASS
exit 0

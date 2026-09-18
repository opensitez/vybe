#!/usr/bin/env bash
# vybe-test: bash/pathname_expansion/question_mark_matches_exactly_one_character
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > ab; : > abc; : > a
out=$(echo a?)
[ "$out" = ab ] || fail "want [ab] got [$out]"
out=$(echo a??)
[ "$out" = abc ] || fail "want [abc] got [$out]"
echo PASS
exit 0

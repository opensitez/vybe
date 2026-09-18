#!/usr/bin/env bash
# vybe-test: bash/pathname_expansion/matches_with_spaces_stay_single_words
# Pathname expansion happens after word splitting, so a matched name is never
# split again.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > "a b.txt"; : > "c d.txt"
[ "$(count *.txt)" = 2 ] || fail "want 2 got $(count *.txt)"
first() { echo "$1"; }
[ "$(first *.txt)" = "a b.txt" ] || fail "got [$(first *.txt)]"
echo PASS
exit 0

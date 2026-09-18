#!/usr/bin/env bash
# vybe-test: bash/pathname_expansion/star_skips_names_starting_with_dot
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > a.txt; : > b.txt; : > .hidden
[ "$(count *)" = 2 ] || fail "want 2 got $(count *)"
echo PASS
exit 0

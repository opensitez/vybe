#!/usr/bin/env bash
# vybe-test: bash/pathname_expansion/results_are_sorted_by_name
# Matches are sorted according to LC_COLLATE; in the C locale digits sort
# before upper case, which sorts before lower case.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
export LC_ALL=C
: > c; : > a; : > B; : > 1
out=$(echo *)
[ "$out" = '1 B a c' ] || fail "want [1 B a c] got [$out]"
echo PASS
exit 0

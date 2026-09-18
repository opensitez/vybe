#!/usr/bin/env bash
# vybe-test: bash/failglob_behavior/matching_pattern_and_literals_are_unaffected
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > a.txt
shopt -s failglob
out=$(echo *.txt plain 2>&1); st=$?
[ "$st" -eq 0 ] && [ "$out" = "a.txt plain" ] || fail "st=$st got [$out]"
echo PASS
exit 0

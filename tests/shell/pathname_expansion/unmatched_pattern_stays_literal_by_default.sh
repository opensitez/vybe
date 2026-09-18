#!/usr/bin/env bash
# vybe-test: bash/pathname_expansion/unmatched_pattern_stays_literal_by_default
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
out=$(echo *.zzz nomatch?)
[ "$out" = '*.zzz nomatch?' ] || fail "got [$out]"
echo PASS
exit 0

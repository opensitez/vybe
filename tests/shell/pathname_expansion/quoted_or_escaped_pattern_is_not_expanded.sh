#!/usr/bin/env bash
# vybe-test: bash/pathname_expansion/quoted_or_escaped_pattern_is_not_expanded
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > a.txt
out=$(echo '*.txt' "*.txt" \*.txt *".txt")
[ "$out" = '*.txt *.txt *.txt a.txt' ] || fail "got [$out]"
echo PASS
exit 0

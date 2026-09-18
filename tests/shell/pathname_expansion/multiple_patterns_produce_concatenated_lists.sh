#!/usr/bin/env bash
# vybe-test: bash/pathname_expansion/multiple_patterns_produce_concatenated_lists
# Each word is expanded on its own; the results keep the order of the words,
# not a global sort.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
export LC_ALL=C
: > a1; : > a2; : > b1
out=$(echo b* a*)
[ "$out" = 'b1 a1 a2' ] || fail "got [$out]"
echo PASS
exit 0

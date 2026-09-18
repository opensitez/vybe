#!/usr/bin/env bash
# vybe-test: bash/expansion_order_and_interactions/parameter_expansion_then_splitting_then_globbing
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
export LC_ALL=C
: > a.txt; : > b.txt
x='*.txt lit'
out=$(echo $x)
[ "$out" = "a.txt b.txt lit" ] || fail "got [$out]"
echo PASS
exit 0

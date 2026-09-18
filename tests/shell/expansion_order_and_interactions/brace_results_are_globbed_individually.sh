#!/usr/bin/env bash
# vybe-test: bash/expansion_order_and_interactions/brace_results_are_globbed_individually
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
export LC_ALL=C
: > a.txt; : > b.txt; : > c.log
out=$(echo *.{txt,log,zzz})
[ "$out" = "a.txt b.txt c.log *.zzz" ] || fail "got [$out]"
echo PASS
exit 0

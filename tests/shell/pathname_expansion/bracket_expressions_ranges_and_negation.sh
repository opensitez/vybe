#!/usr/bin/env bash
# vybe-test: bash/pathname_expansion/bracket_expressions_ranges_and_negation
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
export LC_ALL=C
: > a1; : > b2; : > c3; : > d4
[ "$(echo [ab]?)" = "a1 b2" ] || fail "[ab]? got [$(echo [ab]?)]"
[ "$(echo [!a-b]?)" = "c3 d4" ] || fail "[!a-b]? got [$(echo [!a-b]?)]"
[ "$(echo ?[[:digit:]])" = "a1 b2 c3 d4" ] || fail "class got [$(echo ?[[:digit:]])]"
[ "$(echo ?[3-4])" = "c3 d4" ] || fail "range got [$(echo ?[3-4])]"
echo PASS
exit 0

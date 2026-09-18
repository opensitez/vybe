#!/usr/bin/env bash
# vybe-test: bash/pathname_expansion/slash_is_matched_only_literally
# * and ? never match /, so each directory component needs its own pattern.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
mkdir sub; : > sub/x.txt; : > top.txt
[ "$(echo *.txt)" = top.txt ] || fail "* must not descend: got [$(echo *.txt)]"
[ "$(echo */*.txt)" = sub/x.txt ] || fail "*/*.txt got [$(echo */*.txt)]"
[ "$(echo sub/*)" = sub/x.txt ] || fail "sub/* got [$(echo sub/*)]"
[ "$(echo *x.txt)" = '*x.txt' ] || fail "*x.txt must not match sub/x.txt: got [$(echo *x.txt)]"
echo PASS
exit 0

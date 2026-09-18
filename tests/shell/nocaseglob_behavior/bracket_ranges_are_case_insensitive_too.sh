#!/usr/bin/env bash
# vybe-test: bash/nocaseglob_behavior/bracket_ranges_are_case_insensitive_too
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
export LC_ALL=C
: > B.txt; : > z.txt
shopt -s nocaseglob
[ "$(echo [a-c].txt)" = B.txt ] || fail "got [$(echo [a-c].txt)]"
echo PASS
exit 0

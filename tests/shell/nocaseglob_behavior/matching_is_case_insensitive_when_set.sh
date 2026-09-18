#!/usr/bin/env bash
# vybe-test: bash/nocaseglob_behavior/matching_is_case_insensitive_when_set
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > Readme.TXT
[ "$(echo *.txt)" = '*.txt' ] || fail "default is case sensitive: got [$(echo *.txt)]"
shopt -s nocaseglob
[ "$(echo *.txt)" = Readme.TXT ] || fail "nocaseglob: got [$(echo *.txt)]"
[ "$(echo readme.*)" = Readme.TXT ] || fail "nocaseglob prefix: got [$(echo readme.*)]"
echo PASS
exit 0

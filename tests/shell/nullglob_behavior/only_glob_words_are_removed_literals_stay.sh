#!/usr/bin/env bash
# vybe-test: bash/nullglob_behavior/only_glob_words_are_removed_literals_stay
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
shopt -s nullglob
[ "$(count *.zzz literal '*.zzz')" = 2 ] || fail "want 2 got $(count *.zzz literal '*.zzz')"
echo PASS
exit 0

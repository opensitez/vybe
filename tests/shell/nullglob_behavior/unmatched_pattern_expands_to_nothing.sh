#!/usr/bin/env bash
# vybe-test: bash/nullglob_behavior/unmatched_pattern_expands_to_nothing
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
shopt -s nullglob
[ "$(count *.zzz)" = 0 ] || fail "nullglob on: want 0 got $(count *.zzz)"
shopt -u nullglob
[ "$(count *.zzz)" = 1 ] || fail "nullglob off: want 1 literal word"
echo PASS
exit 0

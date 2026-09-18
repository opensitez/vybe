#!/usr/bin/env bash
# vybe-test: bash/nullglob_behavior/array_from_unmatched_pattern_is_empty
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
shopt -s nullglob
files=(*.zzz)
[ "${#files[@]}" -eq 0 ] || fail "want 0 elements got ${#files[@]}"
: > one.zzz
files=(*.zzz)
[ "${#files[@]}" -eq 1 ] || fail "want 1 element got ${#files[@]}"
echo PASS
exit 0

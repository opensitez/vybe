#!/usr/bin/env bash
# vybe-test: bash/dotglob_behavior/dotglob_lets_star_match_hidden_names
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > a; : > .hidden
[ "$(count *)" = 1 ] || fail "default: want 1 got $(count *)"
shopt -s dotglob
[ "$(count *)" = 2 ] || fail "dotglob: want 2 got $(count *)"
echo PASS
exit 0

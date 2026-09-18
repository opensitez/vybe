#!/usr/bin/env bash
# vybe-test: bash/globstar_behavior/globstar_skips_hidden_unless_dotglob
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
mkdir .hid; : > .hid/x; : > a
shopt -s globstar
[ "$(count **)" = 1 ] || fail "default: want 1 got $(count **)"
shopt -s dotglob
[ "$(count **)" = 3 ] || fail "dotglob: want 3 got $(count **)"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/dotglob_behavior/explicit_leading_dot_always_matches_hidden
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > .hidden; : > .other
[ "$(echo .h*)" = .hidden ] || fail "got [$(echo .h*)]"
echo PASS
exit 0

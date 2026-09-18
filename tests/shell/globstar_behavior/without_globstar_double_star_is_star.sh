#!/usr/bin/env bash
# vybe-test: bash/globstar_behavior/without_globstar_double_star_is_star
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
export LC_ALL=C
mkdir sub; : > a; : > sub/b
shopt -u globstar
out=$(echo **)
[ "$out" = "a sub" ] || fail "got [$out]"
echo PASS
exit 0

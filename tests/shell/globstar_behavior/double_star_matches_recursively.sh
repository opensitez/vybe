#!/usr/bin/env bash
# vybe-test: bash/globstar_behavior/double_star_matches_recursively
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
export LC_ALL=C
mkdir -p sub/deep; : > a; : > sub/b; : > sub/deep/c
shopt -s globstar
out=$(echo **)
[ "$out" = "a sub sub/b sub/deep sub/deep/c" ] || fail "got [$out]"
echo PASS
exit 0

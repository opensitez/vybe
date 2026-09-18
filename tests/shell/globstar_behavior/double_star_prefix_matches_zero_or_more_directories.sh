#!/usr/bin/env bash
# vybe-test: bash/globstar_behavior/double_star_prefix_matches_zero_or_more_directories
# **/*.txt includes top-level *.txt because ** can match no directory at all.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
export LC_ALL=C
mkdir -p sub/deep; : > top.txt; : > sub/mid.txt; : > sub/deep/low.txt; : > sub/other.log
shopt -s globstar
out=$(echo **/*.txt)
[ "$out" = "sub/deep/low.txt sub/mid.txt top.txt" ] || fail "got [$out]"
echo PASS
exit 0

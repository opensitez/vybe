#!/usr/bin/env bash
# vybe-test: bash/globstar_behavior/double_star_is_only_special_as_a_whole_component
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
export LC_ALL=C
mkdir sub; : > a1; : > sub/a2
shopt -s globstar
out=$(echo a**)
[ "$out" = a1 ] || fail "a** must behave like a*: got [$out]"
echo PASS
exit 0

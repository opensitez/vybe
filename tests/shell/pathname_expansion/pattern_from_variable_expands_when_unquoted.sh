#!/usr/bin/env bash
# vybe-test: bash/pathname_expansion/pattern_from_variable_expands_when_unquoted
# The pattern is kept literal in the assignment and expanded only when the
# variable is used unquoted.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > a.txt; : > b.txt
p=*.txt
[ "$p" = '*.txt' ] || fail "assignment must not glob: got [$p]"
[ "$(count $p)" = 2 ] || fail "unquoted use globs: want 2 got $(count $p)"
[ "$(count "$p")" = 1 ] || fail "quoted use stays literal"
echo PASS
exit 0

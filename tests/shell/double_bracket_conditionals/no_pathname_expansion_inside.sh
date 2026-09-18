#!/usr/bin/env bash
# vybe-test: bash/double_bracket_conditionals/no_pathname_expansion_inside
# * inside [[ ]] is a literal file name, not a glob.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > a.txt
[[ -e *.txt ]] && fail "*.txt must be tested literally and not exist"
[ -e *.txt ] || fail "[ ] does expand the glob"
: > '*.txt'
[[ -e *.txt ]] || fail "now the literal name exists"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/file_test_operators/ownership_tests_O_and_G
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > mine
[ -O mine ] || fail "a file I created is owned by me"
[ -G mine ] || fail "a file I created has my group"
[ -O missing ] && fail "-O on missing"
echo PASS
exit 0

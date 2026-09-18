#!/usr/bin/env bash
# vybe-test: bash/file_test_operators/identity_test_ef_compares_inodes
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > a; : > b; ln -s a link; ln a hard
[ a -ef link ] || fail "file and its symlink are the same file"
[ a -ef hard ] || fail "hard links are the same file"
[ a -ef b ] && fail "distinct files"
[ a -ef ./a ] || fail "same path spelled differently"
echo PASS
exit 0

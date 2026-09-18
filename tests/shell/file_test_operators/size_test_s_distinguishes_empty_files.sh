#!/usr/bin/env bash
# vybe-test: bash/file_test_operators/size_test_s_distinguishes_empty_files
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > empty; echo x > nonempty
[ -s nonempty ] || fail "-s on non-empty"
[ -s empty ] && fail "-s on empty"
[ -s missing ] && fail "-s on missing"
echo PASS
exit 0

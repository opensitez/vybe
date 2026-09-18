#!/usr/bin/env bash
# vybe-test: bash/file_test_operators/timestamp_comparison_nt_and_ot
# -nt/-ot compare modification times; a missing operand counts as older.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > old; touch -t 202001010000 old; : > new
[ new -nt old ] || fail "new -nt old"
[ old -ot new ] || fail "old -ot new"
[ old -nt new ] && fail "old -nt new must be false"
[ new -nt new ] && fail "a file is not newer than itself"
[ new -nt missing ] || fail "existing -nt missing is true"
[ missing -ot new ] || fail "missing -ot existing is true"
echo PASS
exit 0

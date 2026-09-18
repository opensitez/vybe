#!/usr/bin/env bash
# vybe-test: bash/file_test_operators/existence_and_type_e_f_d
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > file; mkdir dir
[ -e file ] && [ -e dir ] || fail "-e on file and dir"
[ -e missing ] && fail "-e on missing"
[ -f file ] || fail "-f on file"
[ -f dir ] && fail "-f on dir"
[ -d dir ] || fail "-d on dir"
[ -d file ] && fail "-d on file"
[[ -f file && -d dir && ! -e missing ]] || fail "same operators inside [["
echo PASS
exit 0

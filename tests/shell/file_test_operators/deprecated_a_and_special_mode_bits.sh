#!/usr/bin/env bash
# vybe-test: bash/file_test_operators/deprecated_a_and_special_mode_bits
# -a (unary) is an alias of -e; -k -u -g test sticky, setuid and setgid bits.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > f
[ -a f ] || fail "-a on existing"
[ -a missing ] && fail "-a on missing"
[ -k f ] && fail "sticky on plain file"
[ -u f ] && fail "setuid on plain file"
[ -g f ] && fail "setgid on plain file"
chmod u+s f
[ -u f ] || fail "-u after chmod u+s"
echo PASS
exit 0

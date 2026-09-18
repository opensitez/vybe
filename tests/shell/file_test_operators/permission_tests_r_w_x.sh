#!/usr/bin/env bash
# vybe-test: bash/file_test_operators/permission_tests_r_w_x
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=$(mktemp -d) || fail mktemp; trap 'rm -rf "$tmp"' EXIT; cd "$tmp" || fail cd
: > f
[ -r f ] && [ -w f ] || fail "new file is readable and writable"
[ -x f ] && fail "new file must not be executable"
chmod +x f
[ -x f ] || fail "-x after chmod +x"
[ -x "$tmp" ] || fail "a directory with search permission is -x"
[ -r missing ] && fail "-r on missing"
echo PASS
exit 0

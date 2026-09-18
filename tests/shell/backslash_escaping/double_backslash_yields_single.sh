#!/usr/bin/env bash
# vybe-test: bash/backslash_escaping/double_backslash_yields_single
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=\\
[ "${#x}" -eq 1 ] || fail "want one char got ${#x}"
[ "$x" = '\' ] || fail "want backslash got [$x]"
echo PASS
exit 0

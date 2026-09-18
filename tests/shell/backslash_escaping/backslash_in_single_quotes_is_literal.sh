#!/usr/bin/env bash
# vybe-test: bash/backslash_escaping/backslash_in_single_quotes_is_literal
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
s='\n\\'
[ "${#s}" -eq 4 ] || fail "want 4 chars got ${#s}"
t='\'
[ "$t" = \\ ] || fail "single backslash in single quotes"
echo PASS
exit 0

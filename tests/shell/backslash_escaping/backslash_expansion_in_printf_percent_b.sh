#!/usr/bin/env bash
# vybe-test: bash/backslash_escaping/backslash_expansion_in_printf_percent_b
# The printf format specifier %b expands backslash escape sequences in its argument.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(printf '%b' 'col1\tcol2\n')
expected=$(printf 'col1\tcol2\n')
[ "$out" = "$expected" ] || fail "printf %b: got [$out]"
echo PASS
exit 0

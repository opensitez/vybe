#!/usr/bin/env bash
# vybe-test: bash/backslash_escaping/backslash_escapes_quote_characters
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(echo \"a\" \'b\')
[ "$out" = "\"a\" 'b'" ] || fail "got [$out]"
echo PASS
exit 0

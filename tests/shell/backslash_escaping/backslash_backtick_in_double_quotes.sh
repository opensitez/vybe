#!/usr/bin/env bash
# vybe-test: bash/backslash_escaping/backslash_backtick_in_double_quotes
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out="\`echo no\`"
[ "$out" = '`echo no`' ] || fail "got [$out]"
echo PASS
exit 0

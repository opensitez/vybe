#!/usr/bin/env bash
# vybe-test: bash/backslash_escaping/backslash_escapes_control_operators
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(echo a\;b c\&d e\|f)
[ "$out" = 'a;b c&d e|f' ] || fail "got [$out]"
out=$(echo \#nocomment)
[ "$out" = '#nocomment' ] || fail "escaped hash: got [$out]"
echo PASS
exit 0

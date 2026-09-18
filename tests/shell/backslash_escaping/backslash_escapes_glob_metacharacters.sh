#!/usr/bin/env bash
# vybe-test: bash/backslash_escaping/backslash_escapes_glob_metacharacters
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(echo \* \? \[x\])
[ "$out" = '* ? [x]' ] || fail "got [$out]"
echo PASS
exit 0

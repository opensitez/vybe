#!/usr/bin/env bash
# vybe-test: bash/backslash_escaping/backslash_newline_after_operator
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(true && \
echo joined)
[ "$out" = joined ] || fail "got [$out]"
echo PASS
exit 0

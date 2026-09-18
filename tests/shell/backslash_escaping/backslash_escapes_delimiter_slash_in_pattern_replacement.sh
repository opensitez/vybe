#!/usr/bin/env bash
# vybe-test: bash/backslash_escaping/backslash_escapes_delimiter_slash_in_pattern_replacement
# In ${var//pattern/replacement}, a forward slash in the pattern is escaped with a backslash.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
path="usr/local/bin"
res="${path//\//-}"
[ "$res" = "usr-local-bin" ] || fail "escaped slash in pattern: want 'usr-local-bin', got [$res]"
echo PASS
exit 0

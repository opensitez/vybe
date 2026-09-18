#!/usr/bin/env bash
# vybe-test: bash/backslash_escaping/backslash_before_ordinary_char_is_removed
# Outside quotes a backslash escapes the next character and is then removed,
# even when that character had no special meaning.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(echo \a\b\c)
[ "$out" = abc ] || fail "want [abc] got [$out]"
echo PASS
exit 0

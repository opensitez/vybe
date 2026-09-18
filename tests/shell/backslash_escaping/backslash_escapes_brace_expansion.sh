#!/usr/bin/env bash
# vybe-test: bash/backslash_escaping/backslash_escapes_brace_expansion
# Escaping braces with backslashes suppresses brace expansion and preserves literal characters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- item_\{a,b\}
[ "$#" -eq 1 ] || fail "arg count: want 1, got $#"
[ "$1" = "item_{a,b}" ] || fail "escaped brace: want 'item_{a,b}', got [$1]"
echo PASS
exit 0

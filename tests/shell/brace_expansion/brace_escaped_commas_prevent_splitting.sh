#!/usr/bin/env bash
# vybe-test: bash/brace_expansion/brace_escaped_commas_prevent_splitting
# Escaping a comma with a backslash inside braces treats the comma as literal text rather than a separator.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- {a\,b,c}
[ "$#" -eq 2 ] || fail "escaped comma count: want 2, got $#"
[ "$1" = "a,b" ] || fail "first word mismatch: got [$1]"
[ "$2" = "c" ] || fail "second word mismatch: got [$2]"
echo PASS
exit 0

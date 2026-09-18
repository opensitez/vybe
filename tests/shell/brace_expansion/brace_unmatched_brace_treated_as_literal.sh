#!/usr/bin/env bash
# vybe-test: bash/brace_expansion/brace_unmatched_brace_treated_as_literal
# Unmatched braces or braces without comma or range are preserved as literal characters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- {a,b
[ "$#" -eq 1 ] || fail "unmatched opening count: want 1, got $#"
[ "$1" = "{a,b" ] || fail "unmatched opening brace mismatch: got [$1]"

set -- a,b}
[ "$#" -eq 1 ] || fail "unmatched closing count: want 1, got $#"
[ "$1" = "a,b}" ] || fail "unmatched closing brace mismatch: got [$1]"

set -- {singleton}
[ "$#" -eq 1 ] || fail "singleton count: want 1, got $#"
[ "$1" = "{singleton}" ] || fail "singleton brace without comma mismatch: got [$1]"
echo PASS
exit 0

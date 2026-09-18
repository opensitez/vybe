#!/usr/bin/env bash
# vybe-test: bash/brace_expansion/brace_escaped_braces_treated_as_literal
# Escaping the curly braces with backslashes outputs literal brace characters without expansion.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- \{a,b\}
[ "$#" -eq 1 ] || fail "count: want 1, got $#"
[ "$1" = "{a,b}" ] || fail "escaped braces failed: got [$1]"
echo PASS
exit 0

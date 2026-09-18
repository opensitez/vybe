#!/usr/bin/env bash
# vybe-test: bash/brace_expansion/brace_nested_inner_expansion
# Nested brace expansions evaluate recursively from inside out.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- {a,b{1,2},c}
[ "$#" -eq 4 ] || fail "nested count: want 4, got $#"
[ "$*" = "a b1 b2 c" ] || fail "nested expansion mismatch: got [$*]"
echo PASS
exit 0

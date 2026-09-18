#!/usr/bin/env bash
# vybe-test: bash/grouping_with_braces/brace_group_single_line_syntax
# A brace group written on a single line { a; b; } requires space after '{' and ';' before '}'.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0; y=0
{ x=1; y=2; }
[ "$x" -eq 1 ] || fail "x in single line brace: want 1, got $x"
[ "$y" -eq 2 ] || fail "y in single line brace: want 2, got $y"
echo PASS
exit 0

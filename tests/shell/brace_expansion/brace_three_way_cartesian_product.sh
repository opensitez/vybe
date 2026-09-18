#!/usr/bin/env bash
# vybe-test: bash/brace_expansion/brace_three_way_cartesian_product
# Three consecutive brace expansions produce the full 2x2x2 = 8 element Cartesian combination.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- {0,1}{a,b}{x,y}
[ "$#" -eq 8 ] || fail "three-way product count: want 8, got $#"
expected="0ax 0ay 0bx 0by 1ax 1ay 1bx 1by"
[ "$*" = "$expected" ] || fail "Cartesian expansion mismatch: want [$expected], got [$*]"
echo PASS
exit 0

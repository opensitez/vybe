#!/usr/bin/env bash
# vybe-test: bash/brace_expansion/brace_cartesian_product_multiplication
# Adjacent brace expansions produce a full Cartesian product of words.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- {a,b}{1,2}
[ "$#" -eq 4 ] || fail "Cartesian count: want 4, got $#"
[ "$*" = "a1 a2 b1 b2" ] || fail "Cartesian product mismatch: got [$*]"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/arithmetic_expansion/ternary_is_right_associative
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((1 ? 2 : 3 ? 4 : 5)) -eq 2 ] || fail "got $((1 ? 2 : 3 ? 4 : 5))"
[ $((0 ? 1 : 0 ? 2 : 3)) -eq 3 ] || fail "got $((0 ? 1 : 0 ? 2 : 3))"
[ $((0 ? 1 : 1 ? 2 : 3)) -eq 2 ] || fail "got $((0 ? 1 : 1 ? 2 : 3))"
echo PASS
exit 0

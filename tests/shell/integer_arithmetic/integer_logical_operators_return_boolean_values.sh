#!/usr/bin/env bash
# vybe-test: bash/integer_arithmetic/integer_logical_operators_return_boolean_values
# Arithmetic logical operators normalize to 0/1 and can be composed.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((5 > 3)) -eq 1 ] || fail "5>3 got $((5 > 3))"
[ $((2 > 3)) -eq 0 ] || fail "2>3 got $((2 > 3))"
[ $((0 && 5)) -eq 0 ] || fail "0&&5 got $((0 && 5))"
[ $((0 || 5)) -eq 1 ] || fail "0||5 got $((0 || 5))"
[ $(! 0) -eq 1 ] || fail "!0 got $((!0))"
[ $(! 1) -eq 0 ] || fail "!1 got $((!1))"
echo PASS
exit 0

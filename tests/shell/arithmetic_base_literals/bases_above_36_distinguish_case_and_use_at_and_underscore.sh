#!/usr/bin/env bash
# vybe-test: bash/arithmetic_base_literals/bases_above_36_distinguish_case_and_use_at_and_underscore
# Up to base 36 letters are case-insensitive; above it, lower case is 10-35,
# upper case 36-61, @ is 62 and _ is 63.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((36#Z)) -eq 35 ] && [ $((36#z)) -eq 35 ] || fail "base 36 case-insensitive"
[ $((64#z)) -eq 35 ] || fail "64#z want 35 got $((64#z))"
[ $((64#Z)) -eq 61 ] || fail "64#Z want 61 got $((64#Z))"
[ $((64#@)) -eq 62 ] || fail "64#@ want 62 got $((64#@))"
[ $((64#_)) -eq 63 ] || fail "64#_ want 63 got $((64#_))"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/integer_arithmetic/integer_bit_shifts_are_bidirectional
# Shift operators move bit positions as arithmetic operators with parentheses.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((1 << 5)) -eq 32 ] || fail "1<<5 got $((1 << 5))"
[ $((32 >> 3)) -eq 4 ] || fail "32>>3 got $((32 >> 3))"
[ $((3 << 2 >> 1)) -eq 6 ] || fail "(3<<2)>>1 got $((3 << 2 >> 1))"
echo PASS
exit 0

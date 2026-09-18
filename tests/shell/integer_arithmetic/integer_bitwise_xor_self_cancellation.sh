#!/usr/bin/env bash
# vybe-test: bash/integer_arithmetic/integer_bitwise_xor_self_cancellation
# x ^ x is always 0 for any integer x.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0x5A
a=$((x ^ x))
b=$((0xA ^ 0x5))
[ "$a" -eq 0 ] || fail "x ^ x = $a"
[ "$b" -eq 15 ] || fail "0xA ^ 0x5 = $b"
echo PASS
exit 0

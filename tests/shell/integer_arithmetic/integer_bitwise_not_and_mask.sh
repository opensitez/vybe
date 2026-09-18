#!/usr/bin/env bash
# vybe-test: bash/integer_arithmetic/integer_bitwise_not_and_mask
# Bitwise complement and and-mask behavior matches two's-complement arithmetic.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((~0x00)) -eq -1 ] || fail "~0x00 got $((~0x00))"
[ $(((~0) & 0x5A)) -eq 90 ] || fail "(~0)&0x5A got $(((~0) & 0x5A))"
echo PASS
exit 0

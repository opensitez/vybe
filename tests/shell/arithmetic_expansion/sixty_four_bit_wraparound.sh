#!/usr/bin/env bash
# vybe-test: bash/arithmetic_expansion/sixty_four_bit_wraparound
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
max=9223372036854775807
[ $((max + 1)) -eq -9223372036854775808 ] || fail "max+1 got $((max + 1))"
[ $((-max - 2)) -eq 9223372036854775807 ] || fail "min-1 got $((-max - 2))"
[ $((2**63)) -eq -9223372036854775808 ] || fail "2**63 got $((2**63))"
echo PASS
exit 0

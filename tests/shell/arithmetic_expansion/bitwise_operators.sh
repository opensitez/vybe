#!/usr/bin/env bash
# vybe-test: bash/arithmetic_expansion/bitwise_operators
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out="$((6&3)) $((6|3)) $((6^3)) $((~0)) $((1<<4)) $((256>>3)) $((-8>>1))"
[ "$out" = "2 7 5 -1 16 32 -4" ] || fail "got [$out]"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/arithmetic_base_literals/explicit_base_notation
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out="$((2#101)) $((8#17)) $((16#ff)) $((36#z)) $((3#12))"
[ "$out" = "5 15 255 35 5" ] || fail "got [$out]"
echo PASS
exit 0

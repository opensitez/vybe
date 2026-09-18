#!/usr/bin/env bash
# vybe-test: bash/arithmetic_base_literals/hex_prefix_is_case_insensitive
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out="$((0x1F)) $((0X1f)) $((0xff)) $((0xFF))"
[ "$out" = "31 31 255 255" ] || fail "got [$out]"
echo PASS
exit 0

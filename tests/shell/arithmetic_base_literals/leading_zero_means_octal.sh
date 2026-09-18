#!/usr/bin/env bash
# vybe-test: bash/arithmetic_base_literals/leading_zero_means_octal
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((010)) -eq 8 ] || fail "010 want 8 got $((010))"
[ $((017 + 1)) -eq 16 ] || fail "017+1 want 16 got $((017 + 1))"
[ $((00)) -eq 0 ] || fail "00 want 0"
echo PASS
exit 0

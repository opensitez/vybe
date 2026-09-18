#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_reuse/disable
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fmt='N=%d'
out1=$(printf "$fmt" 8)
fmt+=' V=%d'
out2=$(printf "$fmt" 8 $((8 + 1)))
[ "$out1" = "N=8" ] || fail "first printf mismatch"
[ "$out2" = "N=8 V=$((8 + 1))" ] || fail "second printf mismatch"
echo PASS
exit 0

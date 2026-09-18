#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_reuse/final
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fmt='N=%d'
out1=$(printf "$fmt" 20)
fmt+=' V=%d'
out2=$(printf "$fmt" 20 $((20 + 1)))
[ "$out1" = "N=20" ] || fail "first printf mismatch"
[ "$out2" = "N=20 V=$((20 + 1))" ] || fail "second printf mismatch"
echo PASS
exit 0

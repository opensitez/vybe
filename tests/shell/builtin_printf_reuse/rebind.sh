#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_reuse/rebind
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fmt='N=%d'
out1=$(printf "$fmt" 16)
fmt+=' V=%d'
out2=$(printf "$fmt" 16 $((16 + 1)))
[ "$out1" = "N=16" ] || fail "first printf mismatch"
[ "$out2" = "N=16 V=$((16 + 1))" ] || fail "second printf mismatch"
echo PASS
exit 0

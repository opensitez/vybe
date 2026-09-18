#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_reuse/cleanup
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fmt='N=%d'
out1=$(printf "$fmt" 15)
fmt+=' V=%d'
out2=$(printf "$fmt" 15 $((15 + 1)))
[ "$out1" = "N=15" ] || fail "first printf mismatch"
[ "$out2" = "N=15 V=$((15 + 1))" ] || fail "second printf mismatch"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_reuse/edge
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fmt='N=%d'
out1=$(printf "$fmt" 19)
fmt+=' V=%d'
out2=$(printf "$fmt" 19 $((19 + 1)))
[ "$out1" = "N=19" ] || fail "first printf mismatch"
[ "$out2" = "N=19 V=$((19 + 1))" ] || fail "second printf mismatch"
echo PASS
exit 0

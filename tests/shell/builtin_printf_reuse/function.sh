#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_reuse/function
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fmt='N=%d'
out1=$(printf "$fmt" 2)
fmt+=' V=%d'
out2=$(printf "$fmt" 2 $((2 + 1)))
[ "$out1" = "N=2" ] || fail "first printf mismatch"
[ "$out2" = "N=2 V=$((2 + 1))" ] || fail "second printf mismatch"
echo PASS
exit 0

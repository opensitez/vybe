#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_reuse/basic
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fmt='N=%d'
out1=$(printf "$fmt" 1)
fmt+=' V=%d'
out2=$(printf "$fmt" 1 $((1 + 1)))
[ "$out1" = "N=1" ] || fail "first printf mismatch"
[ "$out2" = "N=1 V=$((1 + 1))" ] || fail "second printf mismatch"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_reuse/stress
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fmt='N=%d'
out1=$(printf "$fmt" 18)
fmt+=' V=%d'
out2=$(printf "$fmt" 18 $((18 + 1)))
[ "$out1" = "N=18" ] || fail "first printf mismatch"
[ "$out2" = "N=18 V=$((18 + 1))" ] || fail "second printf mismatch"
echo PASS
exit 0

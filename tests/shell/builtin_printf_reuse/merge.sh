#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_reuse/merge
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fmt='N=%d'
out1=$(printf "$fmt" 14)
fmt+=' V=%d'
out2=$(printf "$fmt" 14 $((14 + 1)))
[ "$out1" = "N=14" ] || fail "first printf mismatch"
[ "$out2" = "N=14 V=$((14 + 1))" ] || fail "second printf mismatch"
echo PASS
exit 0

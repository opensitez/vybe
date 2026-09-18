#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_reuse/override
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fmt='N=%d'
out1=$(printf "$fmt" 4)
fmt+=' V=%d'
out2=$(printf "$fmt" 4 $((4 + 1)))
[ "$out1" = "N=4" ] || fail "first printf mismatch"
[ "$out2" = "N=4 V=$((4 + 1))" ] || fail "second printf mismatch"
echo PASS
exit 0

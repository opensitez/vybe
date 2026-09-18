#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_reuse/repeat
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fmt='N=%d'
out1=$(printf "$fmt" 12)
fmt+=' V=%d'
out2=$(printf "$fmt" 12 $((12 + 1)))
[ "$out1" = "N=12" ] || fail "first printf mismatch"
[ "$out2" = "N=12 V=$((12 + 1))" ] || fail "second printf mismatch"
echo PASS
exit 0

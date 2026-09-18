#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_reuse/chain
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fmt='N=%d'
out1=$(printf "$fmt" 5)
fmt+=' V=%d'
out2=$(printf "$fmt" 5 $((5 + 1)))
[ "$out1" = "N=5" ] || fail "first printf mismatch"
[ "$out2" = "N=5 V=$((5 + 1))" ] || fail "second printf mismatch"
echo PASS
exit 0

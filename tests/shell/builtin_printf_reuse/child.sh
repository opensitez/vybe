#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_reuse/child
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fmt='N=%d'
out1=$(printf "$fmt" 11)
fmt+=' V=%d'
out2=$(printf "$fmt" 11 $((11 + 1)))
[ "$out1" = "N=11" ] || fail "first printf mismatch"
[ "$out2" = "N=11 V=$((11 + 1))" ] || fail "second printf mismatch"
echo PASS
exit 0

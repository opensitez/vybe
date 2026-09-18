#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_reuse/redefine
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fmt='N=%d'
out1=$(printf "$fmt" 9)
fmt+=' V=%d'
out2=$(printf "$fmt" 9 $((9 + 1)))
[ "$out1" = "N=9" ] || fail "first printf mismatch"
[ "$out2" = "N=9 V=$((9 + 1))" ] || fail "second printf mismatch"
echo PASS
exit 0

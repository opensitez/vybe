#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_reuse/escaped
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fmt='N=%d'
out1=$(printf "$fmt" 7)
fmt+=' V=%d'
out2=$(printf "$fmt" 7 $((7 + 1)))
[ "$out1" = "N=7" ] || fail "first printf mismatch"
[ "$out2" = "N=7 V=$((7 + 1))" ] || fail "second printf mismatch"
echo PASS
exit 0

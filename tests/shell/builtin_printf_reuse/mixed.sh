#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_reuse/mixed
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fmt='N=%d'
out1=$(printf "$fmt" 17)
fmt+=' V=%d'
out2=$(printf "$fmt" 17 $((17 + 1)))
[ "$out1" = "N=17" ] || fail "first printf mismatch"
[ "$out2" = "N=17 V=$((17 + 1))" ] || fail "second printf mismatch"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_reuse/boundary
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fmt='N=%d'
out1=$(printf "$fmt" 13)
fmt+=' V=%d'
out2=$(printf "$fmt" 13 $((13 + 1)))
[ "$out1" = "N=13" ] || fail "first printf mismatch"
[ "$out2" = "N=13 V=$((13 + 1))" ] || fail "second printf mismatch"
echo PASS
exit 0

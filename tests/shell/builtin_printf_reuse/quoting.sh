#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_reuse/quoting
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fmt='N=%d'
out1=$(printf "$fmt" 6)
fmt+=' V=%d'
out2=$(printf "$fmt" 6 $((6 + 1)))
[ "$out1" = "N=6" ] || fail "first printf mismatch"
[ "$out2" = "N=6 V=$((6 + 1))" ] || fail "second printf mismatch"
echo PASS
exit 0

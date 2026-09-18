#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_reuse/subshell
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fmt='N=%d'
out1=$(printf "$fmt" 3)
fmt+=' V=%d'
out2=$(printf "$fmt" 3 $((3 + 1)))
[ "$out1" = "N=3" ] || fail "first printf mismatch"
[ "$out2" = "N=3 V=$((3 + 1))" ] || fail "second printf mismatch"
echo PASS
exit 0

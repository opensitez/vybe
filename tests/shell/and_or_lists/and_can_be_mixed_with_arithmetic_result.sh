#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/and_can_be_mixed_with_arithmetic_result
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
((1)) && x=1 && x=2
[ "$x" -eq 2 ] || fail "arithmetic success should enter and-list"
echo PASS
exit 0

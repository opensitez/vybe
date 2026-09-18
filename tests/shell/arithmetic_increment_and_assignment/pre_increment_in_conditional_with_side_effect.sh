#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/pre_increment_in_conditional_with_side_effect
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
if (( ++x )); then then=ok; else then=bad; fi
[ "$then" = ok ] || fail "++x from 0 should be true"
if (( --x )); then then=bad; else then=ok; fi
[ "$then" = ok ] || fail "--x from 0 should be false"
[ "$x" -eq 0 ] || fail "x should return to 0"
echo PASS
exit 0

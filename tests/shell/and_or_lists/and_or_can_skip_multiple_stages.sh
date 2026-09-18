#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/and_or_can_skip_multiple_stages
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
steps=0
: && steps=$((steps + 1)) && steps=$((steps + 1))
[ "$steps" -eq 2 ] || fail "two-ands should both run after a success"
steps=0
: && false || steps=$((steps + 1))
[ "$steps" -eq 1 ] || fail "false in the middle should skip second and run or"
steps=0
false || : && steps=$((steps + 1))
[ "$steps" -eq 1 ] || fail "true rhs in or should allow and rhs"
echo PASS
exit 0

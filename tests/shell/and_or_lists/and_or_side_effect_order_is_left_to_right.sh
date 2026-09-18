#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/and_or_side_effect_order_is_left_to_right
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
log=""
: && log+="A" || log+="B"
[ "$log" = "A" ] || fail "left side must execute first and set final log"
log=""
false || log+="B" && log+="C"
[ "$log" = "BC" ] || fail "or rhs should run, then following and"
echo PASS
exit 0

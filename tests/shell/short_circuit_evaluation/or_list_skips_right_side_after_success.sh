#!/usr/bin/env bash
# vybe-test: bash/short_circuit_evaluation/or_list_skips_right_side_after_success
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
side=0
true || side=1
[ "$side" -eq 0 ] || fail "right side ran"
false || side=2
[ "$side" -eq 2 ] || fail "right side must run after failure"
echo PASS
exit 0

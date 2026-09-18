#!/usr/bin/env bash
# vybe-test: bash/short_circuit_evaluation/and_list_skips_right_side_after_failure
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
side=0
false && side=1
[ "$side" -eq 0 ] || fail "right side ran"
true && side=2
[ "$side" -eq 2 ] || fail "right side must run after success"
echo PASS
exit 0

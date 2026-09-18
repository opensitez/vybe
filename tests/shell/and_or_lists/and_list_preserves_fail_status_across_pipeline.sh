#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/and_list_preserves_fail_status_across_pipeline
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count=0
( : )
: && ((count += 1))
[ "$count" -eq 1 ] || fail "rhs should still run with subshell lhs success"
count=0
(false) && ((count += 1))
[ "$count" -eq 0 ] || fail "subshelled failure should skip rhs"
echo PASS
exit 0

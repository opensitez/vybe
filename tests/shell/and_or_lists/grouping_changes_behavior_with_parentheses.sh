#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/grouping_changes_behavior_with_parentheses
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
log=""
: && (false || log=from_group)
[ "$log" = "from_group" ] || fail "grouping should allow or after and"

log=""
: || (false && log=from_group)
[ "$log" = "" ] || fail "grouping should preserve short-circuit inside false or"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/child_status_collection/track_children_across_nested_subshells
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
( : & )
p=$!
wait "$p"
[ "$p" -gt 0 ] || fail "subshelled background has pid"
echo PASS
exit 0

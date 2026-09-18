#!/usr/bin/env bash
# vybe-test: bash/child_status_collection/wait_n_captures_any_child
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: &
: &
p=$(wait -n 2>/dev/null)
st=$?
[ "$st" -eq 0 ] || fail "wait -n should succeed"
[ -n "$p" ] || fail "wait -n should return a pid"
wait
wait "$p" 2>/dev/null || true
echo PASS
exit 0

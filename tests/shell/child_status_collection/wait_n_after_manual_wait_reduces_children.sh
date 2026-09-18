#!/usr/bin/env bash
# vybe-test: bash/child_status_collection/wait_n_after_manual_wait_reduces_children
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: &
p1=$!
: &
p2=$!
wait "$p1"
s=0
wait -n || s=$?
wait "$p2"
[ "$s" -eq 0 ] || fail "wait -n should eventually succeed"
echo PASS
exit 0

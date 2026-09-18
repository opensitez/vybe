#!/usr/bin/env bash
# vybe-test: bash/child_status_collection/wait_n_collects_status_of_long_running
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
( : ) &
( : ) &
# first child completion should return quickly via -n
while wait -n; do break; done
s=$?
[ "$s" -eq 0 ] || fail "wait -n one child should succeed"
echo PASS
exit 0

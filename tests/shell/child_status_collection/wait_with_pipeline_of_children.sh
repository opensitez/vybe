#!/usr/bin/env bash
# vybe-test: bash/child_status_collection/wait_with_pipeline_of_children
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(
  :
) &
p=$!
wait "$p"
[ "$?" -eq 0 ] || fail "pipeline child wait"
echo PASS
exit 0

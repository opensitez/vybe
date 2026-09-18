#!/usr/bin/env bash
# vybe-test: bash/child_status_collection/pid_becomes_unset_after_reap
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: &
p=$!
wait "$p"
if wait "$p" 2>/dev/null; then
  fail "reaping twice should fail"
fi
echo PASS
exit 0

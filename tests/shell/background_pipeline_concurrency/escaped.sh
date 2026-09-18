#!/usr/bin/env bash
# vybe-test: bash/background_pipeline_concurrency/escaped
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
pids=()
for ((i=0; i<7; i++)); do
  (exit $((i % 2))) &
  pids+=("$!")
done
idx=0
for pid in "${pids[@]}"; do
  wait "$pid"
  status=$?
  expect=$((idx % 2))
  [ "$status" -eq "$expect" ] || fail "wait status mismatch at $idx"
  idx=$((idx + 1))
done
[ "$idx" -eq 7 ] || fail "wrong number of waits"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/child_status_collection/collect_status_for_many_children
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
pids=()
for ((i=0;i<3;i++)); do : & pids+=("$!"); done
for p in "${pids[@]}"; do wait "$p"; done
echo PASS
exit 0

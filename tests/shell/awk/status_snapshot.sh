#!/usr/bin/env bash
# vybe-test: bash/awk/status_snapshot
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=6
awk_count=0
awk(){ awk_count=$((IDX + 1)); }
awk
(( awk_count == IDX + 1 )) || fail "function named awk did not execute"
unset -f awk
if (( IDX % 2 == 0 )); then
  [[ ${#BASH_VERSION} -ge 1 ]] || fail "awk metadata marker missing"
fi
echo PASS
exit 0

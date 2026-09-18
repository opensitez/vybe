#!/usr/bin/env bash
# vybe-test: bash/associative_array_iteration/boundary
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map
for ((i=0; i<13; i++)); do
  map["k$i"]="v$i"
done
count=0
for _ in "${!map[@]}"; do
  count=$((count + 1))
done
[ "$count" -eq 13 ] || fail "assoc iteration mismatch"
echo PASS
exit 0

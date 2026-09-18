#!/usr/bin/env bash
# vybe-test: bash/associative_array_assignment/boundary
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A src
for ((i=0; i<13; i++)); do
  src["k$i"]="v$i"
done
declare -A dst
for key in "${!src[@]}"; do
  dst["$key"]="${src[$key]}"
done
[ "${#src[@]}" -eq "${#dst[@]}" ] || fail "assoc assignment lost entries"
echo PASS
exit 0

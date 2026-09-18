#!/usr/bin/env bash
# vybe-test: bash/associative_array_assignment/basic
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A src
for ((i=0; i<1; i++)); do
  src["k$i"]="v$i"
done
declare -A dst
for key in "${!src[@]}"; do
  dst["$key"]="${src[$key]}"
done
[ "${#src[@]}" -eq "${#dst[@]}" ] || fail "assoc assignment lost entries"
echo PASS
exit 0

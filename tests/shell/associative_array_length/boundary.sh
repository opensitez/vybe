#!/usr/bin/env bash
# vybe-test: bash/associative_array_length/boundary
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map
for ((i=0; i<13; i++)); do
  map["k$i"]="v$i"
done
[ "${#map[@]}" -eq 13 ] || fail "length operator mismatch"
echo PASS
exit 0

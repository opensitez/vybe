#!/usr/bin/env bash
# vybe-test: bash/associative_array_length/edge
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map
for ((i=0; i<19; i++)); do
  map["k$i"]="v$i"
done
[ "${#map[@]}" -eq 19 ] || fail "length operator mismatch"
echo PASS
exit 0

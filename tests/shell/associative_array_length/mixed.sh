#!/usr/bin/env bash
# vybe-test: bash/associative_array_length/mixed
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map
for ((i=0; i<17; i++)); do
  map["k$i"]="v$i"
done
[ "${#map[@]}" -eq 17 ] || fail "length operator mismatch"
echo PASS
exit 0

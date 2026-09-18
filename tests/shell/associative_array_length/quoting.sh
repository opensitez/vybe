#!/usr/bin/env bash
# vybe-test: bash/associative_array_length/quoting
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map
for ((i=0; i<6; i++)); do
  map["k$i"]="v$i"
done
[ "${#map[@]}" -eq 6 ] || fail "length operator mismatch"
echo PASS
exit 0

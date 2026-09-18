#!/usr/bin/env bash
# vybe-test: bash/associative_array_creation/stress
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map
for ((i=0; i<18; i++)); do
  map["k$i"]="v$i"
done
[ "${#map[@]}" -eq 18 ] || fail "assoc length wrong"
last=$((18 - 1))
[ "${map["k$last"]}" = "v$last" ] || fail "assoc key lookup failed"
echo PASS
exit 0

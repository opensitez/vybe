#!/usr/bin/env bash
# vybe-test: bash/associative_array_creation/escaped
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map
for ((i=0; i<7; i++)); do
  map["k$i"]="v$i"
done
[ "${#map[@]}" -eq 7 ] || fail "assoc length wrong"
last=$((7 - 1))
[ "${map["k$last"]}" = "v$last" ] || fail "assoc key lookup failed"
echo PASS
exit 0

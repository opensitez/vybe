#!/usr/bin/env bash
# vybe-test: bash/associative_array_creation/subshell
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map
for ((i=0; i<3; i++)); do
  map["k$i"]="v$i"
done
[ "${#map[@]}" -eq 3 ] || fail "assoc length wrong"
last=$((3 - 1))
[ "${map["k$last"]}" = "v$last" ] || fail "assoc key lookup failed"
echo PASS
exit 0

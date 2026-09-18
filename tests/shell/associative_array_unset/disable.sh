#!/usr/bin/env bash
# vybe-test: bash/associative_array_unset/disable
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map
for ((i=0; i<8; i++)); do
  map["k$i"]="v$i"
done
key="k$((8 - 1))"
unset map["$key"]
if (( 8 > 1 )); then
  expected=$((8 - 1))
else
  expected=0
fi
[ "${#map[@]}" -eq "$expected" ] || fail "unset changed wrong length"
[ -z "${map[$key]+x}" ] || fail "key still present after unset"
echo PASS
exit 0

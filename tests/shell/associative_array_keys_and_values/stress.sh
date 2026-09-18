#!/usr/bin/env bash
# vybe-test: bash/associative_array_keys_and_values/stress
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map
for ((i=0; i<18; i++)); do
  map["k$i"]="v$i"
done
kcount=0
vcount=0
for _ in "${!map[@]}"; do
  kcount=$((kcount + 1))
done
for _ in "${map[@]}"; do
  vcount=$((vcount + 1))
done
[ "$kcount" -eq "$vcount" ] || fail "keys/values count mismatch"
[ "$kcount" -eq 18 ] || fail "unexpected key total"
echo PASS
exit 0

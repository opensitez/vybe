#!/usr/bin/env bash
# vybe-test: bash/array_serialization/boundary
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=()
for ((i=0; i<13; i++)); do
  arr+=("x$i")
done
ser=$(declare -p arr)
unset arr
eval "$ser"
[ "${#arr[@]}" -eq 13 ] || fail "declare -p roundtrip length mismatch"
echo PASS
exit 0

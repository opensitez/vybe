#!/usr/bin/env bash
# vybe-test: bash/array_serialization/override
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=()
for ((i=0; i<4; i++)); do
  arr+=("x$i")
done
ser=$(declare -p arr)
unset arr
eval "$ser"
[ "${#arr[@]}" -eq 4 ] || fail "declare -p roundtrip length mismatch"
echo PASS
exit 0

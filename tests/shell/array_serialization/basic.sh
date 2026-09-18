#!/usr/bin/env bash
# vybe-test: bash/array_serialization/basic
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=()
for ((i=0; i<1; i++)); do
  arr+=("x$i")
done
ser=$(declare -p arr)
unset arr
eval "$ser"
[ "${#arr[@]}" -eq 1 ] || fail "declare -p roundtrip length mismatch"
echo PASS
exit 0

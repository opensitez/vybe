#!/usr/bin/env bash
# vybe-test: bash/array_copy_and_reference_behavior/rebind
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=()
for ((i=0; i<16; i++)); do
  a+=("$i")
done
b=("${a[@]}")
b[0]=999
[ "${a[0]}" = "0" ] || fail "source array mutated by copy"
[ "${#a[@]}" -eq "${#b[@]}" ] || fail "copy changed length"
[ "${b[0]}" = "999" ] || fail "copied array not writable independently"
echo PASS
exit 0

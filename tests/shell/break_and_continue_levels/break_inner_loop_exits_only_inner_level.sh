#!/usr/bin/env bash
# vybe-test: bash/break_and_continue_levels/break_inner_loop_exits_only_inner_level
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=()
for a in 1 2; do
  for b in 1 2; do
    out+=("$a,$b")
    break
  done
  out+=("after$a")
done
[ "${#out[@]}" -eq 4 ] || fail "got ${#out[@]}"
[ "${out[2]}" = after1 ] || fail "expected after1 in third slot"
[ "${out[3]}" = after2 ] || fail "expected after2 in fourth slot"
echo PASS
exit 0

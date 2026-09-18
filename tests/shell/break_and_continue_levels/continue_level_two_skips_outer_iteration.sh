#!/usr/bin/env bash
# vybe-test: bash/break_and_continue_levels/continue_level_two_skips_outer_iteration
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=()
for i in 1 2 3; do
  for j in 1 2 3; do
    [ "$i" -eq 2 ] && continue 2
    out+=("$i:$j")
  done

done
[ "${#out[@]}" -eq 6 ] || fail "count ${#out[@]}"
[ "${out[2]}" = "1:3" ] || fail "third ${out[2]}"
[ "${out[3]}" = "3:1" ] || fail "fourth ${out[3]}"
echo PASS
exit 0

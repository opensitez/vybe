#!/usr/bin/env bash
# vybe-test: bash/break_and_continue_levels/continue_outer_loop_stops_only_outer_iteration
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=()
for x in 1 2 3; do
  for y in 1 2; do
    out+=("$x:$y")
  done
  [ "$x" -eq 2 ] && continue
  out+=("done:$x")
done
[ "${#out[@]}" -eq 8 ] || fail "count ${#out[@]}"
[ "${out[2]}" = "1:2" ] || fail "third ${out[2]}"
[ "${out[3]}" = "done:1" ] || fail "fourth ${out[3]}"
[ "${out[4]}" = "3:1" ] || fail "fifth ${out[4]}"
echo PASS
exit 0

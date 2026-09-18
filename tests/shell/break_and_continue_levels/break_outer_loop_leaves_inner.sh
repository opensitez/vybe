#!/usr/bin/env bash
# vybe-test: bash/break_and_continue_levels/break_outer_loop_leaves_inner
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
seen=()
for x in 1 2 3; do
  for y in 1 2; do
    seen+=("$x:$y")
  done
  [ "$x" -eq 2 ] && break

done
[ "${#seen[@]}" -eq 4 ] || fail "seen ${#seen[@]}"
[ "${seen[2]}" = "2:1" ] || fail "third ${seen[2]}"
[ "${seen[3]}" = "2:2" ] || fail "fourth ${seen[3]}"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/break_and_continue_levels/continue_inner_skips_only_current_iteration
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
vals=()
for i in 1 2; do
  for j in 1 2 3; do
    [ "$j" -eq 2 ] && continue
    vals+=("$i:$j")
  done
done
[ "${#vals[@]}" -eq 4 ] || fail "count ${#vals[@]}"
[ "${vals[0]}" = "1:1" ] || fail "first ${vals[0]}"
[ "${vals[1]}" = "1:3" ] || fail "second ${vals[1]}"
echo PASS
exit 0

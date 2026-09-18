#!/usr/bin/env bash
# vybe-test: bash/break_and_continue_levels/continue_in_case_inside_loop
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=()
for x in 1 2 3 4; do
  case "$x" in
    2|4)
      continue
      ;;
    *)
      out+=("$x")
      ;;
  esac
done
[ "${#out[@]}" -eq 2 ] || fail "got ${#out[@]}"
[ "${out[0]}" = 1 ] || fail "first ${out[0]}"
[ "${out[1]}" = 3 ] || fail "second ${out[1]}"
echo PASS
exit 0

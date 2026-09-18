#!/usr/bin/env bash
# vybe-test: bash/array_argument_safety/disable
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=()
for ((i=0; i<8; i++)); do
  arr+=("v$i")
done
count_args() {
  c=0
  for _ in "$@"; do
    c=$((c + 1))
  done
  [ "$c" -eq 8 ]
}
if ! count_args "${arr[@]}"; then
  fail "argument count from array expansion mismatch"
fi
echo PASS
exit 0

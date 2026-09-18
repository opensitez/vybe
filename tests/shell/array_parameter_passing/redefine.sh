#!/usr/bin/env bash
# vybe-test: bash/array_parameter_passing/redefine
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
func_count() {
  local c=0
  for _ in "$@"; do
    c=$((c + 1))
  done
  [ "$c" -eq 9 ]
}
arr=()
for ((i=0; i<9; i++)); do
  arr+=("$i")
done
if ! func_count "${arr[@]}"; then
  fail "parameter array passing failed"
fi
echo PASS
exit 0

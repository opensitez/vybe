#!/usr/bin/env bash
# vybe-test: bash/array_parameter_passing/mixed
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
func_count() {
  local c=0
  for _ in "$@"; do
    c=$((c + 1))
  done
  [ "$c" -eq 17 ]
}
arr=()
for ((i=0; i<17; i++)); do
  arr+=("$i")
done
if ! func_count "${arr[@]}"; then
  fail "parameter array passing failed"
fi
echo PASS
exit 0

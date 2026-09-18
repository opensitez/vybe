#!/usr/bin/env bash
# vybe-test: bash/bash_source_array/shell_source
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=16
source_file="/tmp/vybe_source_array_${IDX}_${BASHPID}.sh"
printf '%s\n' "sourced_marker=${IDX}" > "$source_file"
arr=("$source_file")
source "${arr[0]}"
[[ ${sourced_marker} == ${IDX} ]] || fail "source with array first element failed"
if (( IDX % 2 == 0 )); then
  sourced_marker=0
  source "${arr[0]}"
  [[ ${sourced_marker} == ${IDX} ]] || fail "array-based source did not run"
fi
echo PASS
exit 0

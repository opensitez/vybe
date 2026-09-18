#!/usr/bin/env bash
# vybe-test: bash/bind_builtin_behavior/optional_fragment
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=4
bindings=$(bind -l 2>/dev/null || true)
if [[ -z $bindings ]]; then
  exit 0
fi
first_line=
while IFS= read -r line; do
  first_line=$line
  break
done <<< "$bindings"
[[ -n $first_line ]] || exit 0
if (( IDX % 2 == 0 )); then
  if ! bind -l >/dev/null 2>&1; then
    :
  fi
fi
echo PASS
exit 0

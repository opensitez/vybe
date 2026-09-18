#!/usr/bin/env bash
# vybe-test: bash/bind_keymaps/regex_capture
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=14
if ! bind -l >/dev/null 2>&1; then
  exit 0
fi
maps=$(bind -l)
if [[ -z $maps ]]; then
  exit 0
fi
if (( IDX % 2 == 0 )); then
  names_count=0
  while IFS= read -r line; do
    [[ -n $line ]] && ((names_count += 1))
  done <<< "$maps"
  (( names_count >= 1 )) || fail "no keymaps listed"
fi
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/bind_readline_macros/subshell_probe
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=17
if ! bind -l >/dev/null 2>&1; then
  exit 0
fi
bind '"\C-b": ""'
if (( IDX % 2 == 0 )); then
  if ! bind -P >/dev/null 2>&1; then
    fail "bind -P should be queryable"
  fi
fi
bind -r '\C-b'
echo PASS
exit 0

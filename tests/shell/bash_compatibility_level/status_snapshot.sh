#!/usr/bin/env bash
# vybe-test: bash/bash_compatibility_level/status_snapshot
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=6
if ! shopt -q compat31 >/dev/null 2>&1; then
  [[ -n $BASH_VERSION ]] || fail "no bash version metadata"
else
  shopt -q compat31 >/dev/null 2>&1
  before=$?
  shopt -s compat31
  shopt -q compat31 >/dev/null 2>&1
  after=$?
  (( after == 0 )) || fail "compat31 should be on after shopt -s"
  if (( IDX % 2 == 0 )); then
    shopt -u compat31
    shopt -q compat31 >/dev/null 2>&1
    off=$?
    [[ $off -ne 0 ]] || fail "compat31 should disable"
    shopt -s compat31
  fi
fi
echo PASS
exit 0

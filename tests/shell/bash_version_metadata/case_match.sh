#!/usr/bin/env bash
# vybe-test: bash/bash_version_metadata/case_match
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=13
(( ${#BASH_VERSINFO[@]} >= 5 )) || fail "BASH_VERSINFO too short"
[[ -n ${BASH_VERSINFO[0]} && -n ${BASH_VERSINFO[1]} && -n ${BASH_VERSINFO[2]} ]] || fail "version tuple missing"
[[ -n ${BASH_VERSINFO[4]} ]] || fail "release field missing"
if (( IDX % 2 == 0 )); then
  status=${BASH_VERSINFO[0]}
  (( status >= 0 )) || fail "major must be non-negative"
fi
echo PASS
exit 0

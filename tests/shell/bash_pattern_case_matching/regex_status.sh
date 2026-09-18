#!/usr/bin/env bash
# vybe-test: bash/bash_pattern_case_matching/regex_status
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=15
shopt -s nocasematch
if (( IDX % 2 == 0 )); then
  [[ "TeSt${IDX}" == "test${IDX}" ]] || fail "nocasematch should ignore case"
else
  [[ "vAlUe${IDX}" == "value${IDX}" ]] || fail "nocasematch should ignore case"
fi
shopt -u nocasematch
[[ ! "Case${IDX}" == "case${IDX}" ]] || fail "normal matching should remain case sensitive"
echo PASS
exit 0

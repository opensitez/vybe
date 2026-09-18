#!/usr/bin/env bash
# vybe-test: bash/bash_pattern_regex_special_variables/baseline_validation
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=1
if [[ "X_${IDX}_Y" =~ X_([0-9]+)_Y ]]; then
  (( ${#BASH_REMATCH[@]} >= 2 )) || fail "BASH_REMATCH should expose captures"
  [[ ${BASH_REMATCH[0]} == "X_${IDX}_Y" ]] || fail "BASH_REMATCH[0] mismatch"
  [[ ${BASH_REMATCH[1]} == ${IDX} ]] || fail "BASH_REMATCH[1] mismatch"
else
  fail "regex did not match"
fi
echo PASS
exit 0

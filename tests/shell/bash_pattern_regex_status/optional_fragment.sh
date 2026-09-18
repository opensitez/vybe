#!/usr/bin/env bash
# vybe-test: bash/bash_pattern_regex_status/optional_fragment
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=4
if (( IDX % 2 == 0 )); then
  [[ "abc${IDX}" =~ abc ]] || fail "expected regex hit"
else
  [[ "abc${IDX}" =~ z ]] && fail "unexpected regex hit"
fi
if [[ "abc${IDX}" =~ c$ ]]; then :; else fail "suffix check failed"; fi
echo PASS
exit 0

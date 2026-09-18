#!/usr/bin/env bash
# vybe-test: bash/bash_pattern_extglob_repetition/regex_status
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=15
shopt -s extglob
if (( IDX % 2 == 0 )); then
  [[ "aaaa" == a+(a) ]] || fail "repetition plus should match repeated a"
  [[ "a" == a*(a) ]] || fail "zero-or-more repetition should accept single token"
else
  [[ "b" == a+(a) ]] && fail "repetition should not match base mismatch"
fi
shopt -u extglob
echo PASS
exit 0

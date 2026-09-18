#!/usr/bin/env bash
# vybe-test: bash/bash_pattern_extglob_alternatives/exact_matching
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=3
shopt -s extglob
if (( IDX % 3 == 0 )); then
  [[ "foo" == '@(foo|bar|baz)' ]] || fail "alternation mismatch"
  [[ "qux" == @(foo|bar|baz) ]] && fail "invalid alternation should not match"
elif (( IDX % 3 == 1 )); then
  [[ "bar" == @(foo|bar) ]] || fail "alternation mismatch"
else
  [[ "baz" == @(foo|bar|baz) ]] || fail "alternation mismatch"
fi
shopt -u extglob
echo PASS
exit 0

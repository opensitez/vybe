#!/usr/bin/env bash
# vybe-test: bash/bash_pattern_extglob_grouping/optional_fragment
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=4
shopt -s extglob
if (( IDX % 2 == 0 )); then
  [[ "foobar" == fo@(bar|baz) ]] || fail "grouped alternation should match"
  [[ ! "fooqux" == fo@(bar|baz) ]] || fail "grouped alternation should reject"
else
  [[ "foobaz" == fo@(bar|baz) ]] || fail "grouped alternation should match"
fi
shopt -u extglob
echo PASS
exit 0

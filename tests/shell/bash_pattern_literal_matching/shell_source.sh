#!/usr/bin/env bash
# vybe-test: bash/bash_pattern_literal_matching/shell_source
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=16
[[ 'a*b' == 'a\*b' ]] || fail "escaped asterisk should match literal"
[[ '[xy]' == '\[xy\]' ]] || fail "escaped brackets should match literal"
if (( IDX % 2 == 0 )); then
  [[ '\$x' == '\$x' ]] || fail "escaped dollar should match literal"
fi
echo PASS
exit 0

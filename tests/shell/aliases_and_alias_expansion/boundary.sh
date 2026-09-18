#!/usr/bin/env bash
# vybe-test: bash/aliases_and_alias_expansion/boundary
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
x=0
alias alias_base13='x=$((x + 13))'
alias alias_chain13='alias_base13; x=$((x + 1))'
if (( 13 % 2 == 0 )); then
  alias_chain13
  expected=$((13 + 1))
else
  alias_base13
  expected=13
fi
[ "$x" -eq "$expected" ] || fail "alias expansion mismatch"
if ! alias alias_chain13 >/dev/null 2>&1; then
  fail "alias_chain should exist"
fi
echo PASS
exit 0

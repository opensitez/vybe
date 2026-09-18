#!/usr/bin/env bash
# vybe-test: bash/aliases_and_alias_expansion/parent
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
x=0
alias alias_base10='x=$((x + 10))'
alias alias_chain10='alias_base10; x=$((x + 1))'
if (( 10 % 2 == 0 )); then
  alias_chain10
  expected=$((10 + 1))
else
  alias_base10
  expected=10
fi
[ "$x" -eq "$expected" ] || fail "alias expansion mismatch"
if ! alias alias_chain10 >/dev/null 2>&1; then
  fail "alias_chain should exist"
fi
echo PASS
exit 0

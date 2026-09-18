#!/usr/bin/env bash
# vybe-test: bash/aliases_and_alias_expansion/function
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
x=0
alias alias_base2='x=$((x + 2))'
alias alias_chain2='alias_base2; x=$((x + 1))'
if (( 2 % 2 == 0 )); then
  alias_chain2
  expected=$((2 + 1))
else
  alias_base2
  expected=2
fi
[ "$x" -eq "$expected" ] || fail "alias expansion mismatch"
if ! alias alias_chain2 >/dev/null 2>&1; then
  fail "alias_chain should exist"
fi
echo PASS
exit 0

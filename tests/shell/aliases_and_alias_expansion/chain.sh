#!/usr/bin/env bash
# vybe-test: bash/aliases_and_alias_expansion/chain
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
x=0
alias alias_base5='x=$((x + 5))'
alias alias_chain5='alias_base5; x=$((x + 1))'
if (( 5 % 2 == 0 )); then
  alias_chain5
  expected=$((5 + 1))
else
  alias_base5
  expected=5
fi
[ "$x" -eq "$expected" ] || fail "alias expansion mismatch"
if ! alias alias_chain5 >/dev/null 2>&1; then
  fail "alias_chain should exist"
fi
echo PASS
exit 0

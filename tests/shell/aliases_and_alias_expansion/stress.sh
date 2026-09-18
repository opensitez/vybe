#!/usr/bin/env bash
# vybe-test: bash/aliases_and_alias_expansion/stress
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
x=0
alias alias_base18='x=$((x + 18))'
alias alias_chain18='alias_base18; x=$((x + 1))'
if (( 18 % 2 == 0 )); then
  alias_chain18
  expected=$((18 + 1))
else
  alias_base18
  expected=18
fi
[ "$x" -eq "$expected" ] || fail "alias expansion mismatch"
if ! alias alias_chain18 >/dev/null 2>&1; then
  fail "alias_chain should exist"
fi
echo PASS
exit 0

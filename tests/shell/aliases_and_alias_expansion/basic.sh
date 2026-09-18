#!/usr/bin/env bash
# vybe-test: bash/aliases_and_alias_expansion/basic
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
x=0
alias alias_base1='x=$((x + 1))'
alias alias_chain1='alias_base1; x=$((x + 1))'
if (( 1 % 2 == 0 )); then
  alias_chain1
  expected=$((1 + 1))
else
  alias_base1
  expected=1
fi
[ "$x" -eq "$expected" ] || fail "alias expansion mismatch"
if ! alias alias_chain1 >/dev/null 2>&1; then
  fail "alias_chain should exist"
fi
echo PASS
exit 0

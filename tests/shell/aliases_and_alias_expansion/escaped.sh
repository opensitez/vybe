#!/usr/bin/env bash
# vybe-test: bash/aliases_and_alias_expansion/escaped
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
x=0
alias alias_base7='x=$((x + 7))'
alias alias_chain7='alias_base7; x=$((x + 1))'
if (( 7 % 2 == 0 )); then
  alias_chain7
  expected=$((7 + 1))
else
  alias_base7
  expected=7
fi
[ "$x" -eq "$expected" ] || fail "alias expansion mismatch"
if ! alias alias_chain7 >/dev/null 2>&1; then
  fail "alias_chain should exist"
fi
echo PASS
exit 0

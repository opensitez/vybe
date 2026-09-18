#!/usr/bin/env bash
# vybe-test: bash/aliases_and_alias_expansion/child
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
x=0
alias alias_base11='x=$((x + 11))'
alias alias_chain11='alias_base11; x=$((x + 1))'
if (( 11 % 2 == 0 )); then
  alias_chain11
  expected=$((11 + 1))
else
  alias_base11
  expected=11
fi
[ "$x" -eq "$expected" ] || fail "alias expansion mismatch"
if ! alias alias_chain11 >/dev/null 2>&1; then
  fail "alias_chain should exist"
fi
echo PASS
exit 0

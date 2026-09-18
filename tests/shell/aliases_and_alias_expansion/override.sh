#!/usr/bin/env bash
# vybe-test: bash/aliases_and_alias_expansion/override
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
x=0
alias alias_base4='x=$((x + 4))'
alias alias_chain4='alias_base4; x=$((x + 1))'
if (( 4 % 2 == 0 )); then
  alias_chain4
  expected=$((4 + 1))
else
  alias_base4
  expected=4
fi
[ "$x" -eq "$expected" ] || fail "alias expansion mismatch"
if ! alias alias_chain4 >/dev/null 2>&1; then
  fail "alias_chain should exist"
fi
echo PASS
exit 0

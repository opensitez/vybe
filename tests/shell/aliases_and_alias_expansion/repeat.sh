#!/usr/bin/env bash
# vybe-test: bash/aliases_and_alias_expansion/repeat
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
x=0
alias alias_base12='x=$((x + 12))'
alias alias_chain12='alias_base12; x=$((x + 1))'
if (( 12 % 2 == 0 )); then
  alias_chain12
  expected=$((12 + 1))
else
  alias_base12
  expected=12
fi
[ "$x" -eq "$expected" ] || fail "alias expansion mismatch"
if ! alias alias_chain12 >/dev/null 2>&1; then
  fail "alias_chain should exist"
fi
echo PASS
exit 0

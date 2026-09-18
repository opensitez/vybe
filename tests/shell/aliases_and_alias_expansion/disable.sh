#!/usr/bin/env bash
# vybe-test: bash/aliases_and_alias_expansion/disable
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
x=0
alias alias_base8='x=$((x + 8))'
alias alias_chain8='alias_base8; x=$((x + 1))'
if (( 8 % 2 == 0 )); then
  alias_chain8
  expected=$((8 + 1))
else
  alias_base8
  expected=8
fi
[ "$x" -eq "$expected" ] || fail "alias expansion mismatch"
if ! alias alias_chain8 >/dev/null 2>&1; then
  fail "alias_chain should exist"
fi
echo PASS
exit 0

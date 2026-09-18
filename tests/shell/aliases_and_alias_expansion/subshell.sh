#!/usr/bin/env bash
# vybe-test: bash/aliases_and_alias_expansion/subshell
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
x=0
alias alias_base3='x=$((x + 3))'
alias alias_chain3='alias_base3; x=$((x + 1))'
if (( 3 % 2 == 0 )); then
  alias_chain3
  expected=$((3 + 1))
else
  alias_base3
  expected=3
fi
[ "$x" -eq "$expected" ] || fail "alias expansion mismatch"
if ! alias alias_chain3 >/dev/null 2>&1; then
  fail "alias_chain should exist"
fi
echo PASS
exit 0

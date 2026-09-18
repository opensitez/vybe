#!/usr/bin/env bash
# vybe-test: bash/aliases_and_alias_expansion/merge
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
x=0
alias alias_base14='x=$((x + 14))'
alias alias_chain14='alias_base14; x=$((x + 1))'
if (( 14 % 2 == 0 )); then
  alias_chain14
  expected=$((14 + 1))
else
  alias_base14
  expected=14
fi
[ "$x" -eq "$expected" ] || fail "alias expansion mismatch"
if ! alias alias_chain14 >/dev/null 2>&1; then
  fail "alias_chain should exist"
fi
echo PASS
exit 0

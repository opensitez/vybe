#!/usr/bin/env bash
# vybe-test: bash/aliases_and_alias_expansion/edge
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
x=0
alias alias_base19='x=$((x + 19))'
alias alias_chain19='alias_base19; x=$((x + 1))'
if (( 19 % 2 == 0 )); then
  alias_chain19
  expected=$((19 + 1))
else
  alias_base19
  expected=19
fi
[ "$x" -eq "$expected" ] || fail "alias expansion mismatch"
if ! alias alias_chain19 >/dev/null 2>&1; then
  fail "alias_chain should exist"
fi
echo PASS
exit 0

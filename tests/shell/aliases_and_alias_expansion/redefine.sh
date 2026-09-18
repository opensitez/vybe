#!/usr/bin/env bash
# vybe-test: bash/aliases_and_alias_expansion/redefine
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
x=0
alias alias_base9='x=$((x + 9))'
alias alias_chain9='alias_base9; x=$((x + 1))'
if (( 9 % 2 == 0 )); then
  alias_chain9
  expected=$((9 + 1))
else
  alias_base9
  expected=9
fi
[ "$x" -eq "$expected" ] || fail "alias expansion mismatch"
if ! alias alias_chain9 >/dev/null 2>&1; then
  fail "alias_chain should exist"
fi
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/aliases_and_alias_expansion/final
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
x=0
alias alias_base20='x=$((x + 20))'
alias alias_chain20='alias_base20; x=$((x + 1))'
if (( 20 % 2 == 0 )); then
  alias_chain20
  expected=$((20 + 1))
else
  alias_base20
  expected=20
fi
[ "$x" -eq "$expected" ] || fail "alias expansion mismatch"
if ! alias alias_chain20 >/dev/null 2>&1; then
  fail "alias_chain should exist"
fi
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/aliases_and_alias_expansion/quoting
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
x=0
alias alias_base6='x=$((x + 6))'
alias alias_chain6='alias_base6; x=$((x + 1))'
if (( 6 % 2 == 0 )); then
  alias_chain6
  expected=$((6 + 1))
else
  alias_base6
  expected=6
fi
[ "$x" -eq "$expected" ] || fail "alias expansion mismatch"
if ! alias alias_chain6 >/dev/null 2>&1; then
  fail "alias_chain should exist"
fi
echo PASS
exit 0

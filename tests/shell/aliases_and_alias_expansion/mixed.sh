#!/usr/bin/env bash
# vybe-test: bash/aliases_and_alias_expansion/mixed
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
x=0
alias alias_base17='x=$((x + 17))'
alias alias_chain17='alias_base17; x=$((x + 1))'
if (( 17 % 2 == 0 )); then
  alias_chain17
  expected=$((17 + 1))
else
  alias_base17
  expected=17
fi
[ "$x" -eq "$expected" ] || fail "alias expansion mismatch"
if ! alias alias_chain17 >/dev/null 2>&1; then
  fail "alias_chain should exist"
fi
echo PASS
exit 0

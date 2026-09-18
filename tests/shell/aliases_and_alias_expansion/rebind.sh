#!/usr/bin/env bash
# vybe-test: bash/aliases_and_alias_expansion/rebind
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
x=0
alias alias_base16='x=$((x + 16))'
alias alias_chain16='alias_base16; x=$((x + 1))'
if (( 16 % 2 == 0 )); then
  alias_chain16
  expected=$((16 + 1))
else
  alias_base16
  expected=16
fi
[ "$x" -eq "$expected" ] || fail "alias expansion mismatch"
if ! alias alias_chain16 >/dev/null 2>&1; then
  fail "alias_chain should exist"
fi
echo PASS
exit 0

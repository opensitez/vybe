#!/usr/bin/env bash
# vybe-test: bash/alias_scope_and_lifetime/edge
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
x=0
if (( 19 % 2 == 0 )); then
  f_scope(){ alias scope_alias='x=$((x + 19))'; }
  f_scope
else
  alias scope_alias='x=$((x + 19))'
fi
if (( 19 % 3 == 0 )); then
  scope_alias
else
  (scope_alias)
fi
if (( 19 % 6 == 0 )); then
  expected=19
else
  expected=0
fi
[ "$x" -eq "$expected" ] || fail "unexpected scope result"
if ! alias scope_alias >/dev/null 2>&1; then
  fail "alias should remain defined"
fi
echo PASS
exit 0

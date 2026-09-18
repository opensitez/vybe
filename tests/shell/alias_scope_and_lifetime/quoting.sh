#!/usr/bin/env bash
# vybe-test: bash/alias_scope_and_lifetime/quoting
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
x=0
if (( 6 % 2 == 0 )); then
  f_scope(){ alias scope_alias='x=$((x + 6))'; }
  f_scope
else
  alias scope_alias='x=$((x + 6))'
fi
if (( 6 % 3 == 0 )); then
  scope_alias
else
  (scope_alias)
fi
if (( 6 % 6 == 0 )); then
  expected=6
else
  expected=0
fi
[ "$x" -eq "$expected" ] || fail "unexpected scope result"
if ! alias scope_alias >/dev/null 2>&1; then
  fail "alias should remain defined"
fi
echo PASS
exit 0

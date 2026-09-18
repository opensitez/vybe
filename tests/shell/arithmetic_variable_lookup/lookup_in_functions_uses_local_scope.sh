#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/lookup_in_functions_uses_local_scope
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
f() {
  local x=11
  [ $((x+1)) -eq 12 ] || fail "local x should be 12"
}
x=3
f
[ $((x+1)) -eq 4 ] || fail "global x should remain 3"
echo PASS
exit 0

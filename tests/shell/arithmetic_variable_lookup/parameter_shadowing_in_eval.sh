#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/parameter_shadowing_in_eval
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
outer=5
f() {
  local outer=2
  [ $((outer)) -eq 2 ] || fail "local scope wrong"
  [ $((shadow)) -eq 0 ] || fail "unset in function should be 0"
}
f
[ $((outer)) -eq 5 ] || fail "global should remain 5"
echo PASS
exit 0

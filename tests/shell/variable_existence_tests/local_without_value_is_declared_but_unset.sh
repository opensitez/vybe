#!/usr/bin/env bash
# vybe-test: bash/variable_existence_tests/local_without_value_is_declared_but_unset
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
f() {
  local l
  [[ -v l ]] && fail "local l without value must be unset"
  [ -z "${l+set}" ] || fail "\${l+set} agrees it is unset"
  local m=
  [[ -v m ]] || fail "local m= is set to empty"
}
f
echo PASS
exit 0

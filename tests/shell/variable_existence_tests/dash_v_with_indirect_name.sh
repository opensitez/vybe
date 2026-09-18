#!/usr/bin/env bash
# vybe-test: bash/variable_existence_tests/dash_v_with_indirect_name
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
target=1; unset gone
name=target
[[ -v $name ]] || fail "\$name expands to a set variable"
name=gone
[[ -v $name ]] && fail "\$name expands to an unset variable"
[[ -v name ]] || fail "name itself is set"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/parentheses_and_nested_commands/parentheses_extglob_negation_operator
# In extended globbing, !(pattern) matches anything that does not match the specified patterns.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s extglob
[[ "hello" == !(world) ]] || fail "!(world) should match 'hello'"
[[ "world" == !(world) ]] && fail "!(world) should not match 'world'"
echo PASS
exit 0

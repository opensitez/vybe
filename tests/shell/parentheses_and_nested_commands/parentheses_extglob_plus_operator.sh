#!/usr/bin/env bash
# vybe-test: bash/parentheses_and_nested_commands/parentheses_extglob_plus_operator
# In extended globbing, +(pattern) matches one or more occurrences of the specified patterns.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s extglob
[[ "aaa" == +(a) ]] || fail "+(a) should match 'aaa'"
[[ "ababa" == +(ab)a ]] || fail "+(ab)a should match 'ababa'"
[[ "b" == +(a) ]] && fail "+(a) should not match 'b'"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/parentheses_and_nested_commands/parentheses_extglob_star_operator
# In extended globbing, *(pattern) matches zero or more occurrences of the specified patterns.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s extglob
[[ "" == *(a) ]] || fail "*(a) should match empty string"
[[ "aaaa" == *(a) ]] || fail "*(a) should match 'aaaa'"
[[ "b" == *(a) ]] && fail "*(a) should not match 'b'"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/conditional_expression_syntax/cond_string_lexicographical_greater_than
# The '>' operator tests lexicographical sorting order inside [[ ... ]] according to locale.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ "xyz" > "abc" ]] || fail "'xyz' should be > 'abc'"
[[ "abc" > "xyz" ]] && fail "'abc' should not be > 'xyz'"
echo PASS
exit 0

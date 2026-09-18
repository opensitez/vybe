#!/usr/bin/env bash
# vybe-test: bash/conditional_expression_syntax/cond_string_lexicographical_less_than
# The '<' operator tests lexicographical sorting order inside [[ ... ]] according to locale.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ "abc" < "abd" ]] || fail "'abc' should be < 'abd'"
[[ "abd" < "abc" ]] && fail "'abd' should not be < 'abc'"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/conditional_expression_syntax/cond_string_inequality_not_equal
# The '!=' operator tests string inequality (and non-matching patterns) inside [[ ... ]].
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ "apple" != "orange" ]] || fail "'apple' should not equal 'orange'"
[[ "apple" != "apple" ]] && fail "'apple' != 'apple' should return false"
echo PASS
exit 0

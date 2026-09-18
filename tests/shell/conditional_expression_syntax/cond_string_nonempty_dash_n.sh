#!/usr/bin/env bash
# vybe-test: bash/conditional_expression_syntax/cond_string_nonempty_dash_n
# The -n operator tests whether a string has non-zero length.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
str="content"
[[ -n "$str" ]] || fail "str should be non-empty"
[[ -n "" ]] && fail "empty string should not be non-empty"
echo PASS
exit 0

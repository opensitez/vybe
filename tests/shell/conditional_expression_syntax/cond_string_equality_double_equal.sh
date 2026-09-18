#!/usr/bin/env bash
# vybe-test: bash/conditional_expression_syntax/cond_string_equality_double_equal
# The '==' operator tests string equality (and pattern matching) inside [[ ... ]].
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
s1="alpha"
s2="alpha"
s3="beta"
[[ "$s1" == "$s2" ]] || fail "identical strings should be equal with =="
[[ "$s1" == "$s3" ]] && fail "differing strings should not be equal with =="
echo PASS
exit 0

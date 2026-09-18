#!/usr/bin/env bash
# vybe-test: bash/conditional_expression_syntax/cond_regex_match_operator
# The '=~' binary operator matches a string against an extended regular expression.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ "user@example.com" =~ ^[a-z]+@[a-z]+\.[a-z]+$ ]] || fail "valid email should match regex"
[[ "not_an_email" =~ ^[a-z]+@[a-z]+\.[a-z]+$ ]] && fail "invalid string should not match regex"
echo PASS
exit 0

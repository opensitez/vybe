#!/usr/bin/env bash
# vybe-test: bash/conditional_expression_syntax/cond_string_equality_single_equal
# The single '=' operator is a synonym for '==' in [[ ... ]] conditional expressions.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ "sample" = "sample" ]] || fail "identical strings should be equal with ="
[[ "sample" = "other" ]] && fail "differing strings should not be equal with ="
echo PASS
exit 0

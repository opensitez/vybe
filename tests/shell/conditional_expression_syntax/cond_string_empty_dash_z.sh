#!/usr/bin/env bash
# vybe-test: bash/conditional_expression_syntax/cond_string_empty_dash_z
# The -z operator tests whether a string has zero length.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
empty_val=""
non_empty_val="not_empty"
[[ -z "$empty_val" ]] || fail "empty_val should be zero length"
[[ -z "$non_empty_val" ]] && fail "non_empty_val should not be zero length"
echo PASS
exit 0

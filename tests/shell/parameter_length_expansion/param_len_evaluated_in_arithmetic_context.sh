#!/usr/bin/env bash
# vybe-test: bash/parameter_length_expansion/param_len_evaluated_in_arithmetic_context
# The ${#var} expansion can be embedded directly in arithmetic context (( ... )) calculations.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
token="12345"
(( len_doubled = ${#token} * 2 ))
[ "$len_doubled" -eq 10 ] || fail "arithmetic evaluation failed: want 10, got $len_doubled"
[ $(( ${#token} + 5 )) -eq 10 ] || fail "inline arithmetic with length failed"
echo PASS
exit 0

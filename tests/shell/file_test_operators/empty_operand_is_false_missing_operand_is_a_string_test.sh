#!/usr/bin/env bash
# vybe-test: bash/file_test_operators/empty_operand_is_false_missing_operand_is_a_string_test
# [ -e "" ] tests an empty name (false); [ -e ] has one argument, which is
# the non-empty string "-e" (true).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ -e "" ] && fail "empty name"
[ -d "" ] && fail "empty name -d"
[ -e ] || fail "single argument is a non-empty string test"
echo PASS
exit 0

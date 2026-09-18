#!/usr/bin/env bash
# vybe-test: bash/integer_arithmetic/integer_unset_and_empty_variables_default_to_zero
# Arithmetic treats unset/empty symbols as 0.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset unset_var
empty_var=""
[ $((unset_var + 1)) -eq 1 ] || fail "unset var gave $((unset_var + 1))"
[ $((empty_var + 1)) -eq 1 ] || fail "empty var gave $((empty_var + 1))"
echo PASS
exit 0

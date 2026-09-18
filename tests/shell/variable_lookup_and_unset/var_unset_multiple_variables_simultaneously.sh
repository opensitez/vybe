#!/usr/bin/env bash
# vybe-test: bash/variable_lookup_and_unset/var_unset_multiple_variables_simultaneously
# Passing multiple variable names to unset removes all specified variables in one call.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x1="a" x2="b" x3="c"
unset x1 x2 x3
[[ ! -v x1 ]] || fail "x1 not unset"
[[ ! -v x2 ]] || fail "x2 not unset"
[[ ! -v x3 ]] || fail "x3 not unset"
echo PASS
exit 0

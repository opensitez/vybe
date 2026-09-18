#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/or_runs_rhs_on_lhs_failure
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
false || x=$((x + 1))
[ "$x" -eq 1 ] || fail "rhs should run when lhs fails"
echo PASS
exit 0

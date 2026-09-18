#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/or_skips_rhs_when_lhs_succeeds
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
: || x=$((x + 1))
[ "$x" -eq 0 ] || fail "rhs should be skipped when lhs succeeds"
echo PASS
exit 0

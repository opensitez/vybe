#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/and_executes_rhs_only_on_success
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
: && x=$((x + 1))
[ "$x" -eq 1 ] || fail "rhs should run after success"
echo PASS
exit 0

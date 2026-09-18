#!/usr/bin/env bash
# vybe-test: bash/arithmetic_operator_precedence/logical_and_above_logical_or
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((1 || 0 && 0)) -eq 1 ] || fail "got $((1 || 0 && 0))"
[ $(((1 || 0) && 0)) -eq 0 ] || fail "grouped got $(((1 || 0) && 0))"
echo PASS
exit 0

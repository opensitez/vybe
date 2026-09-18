#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/short_circuiting_in_arithmetic_conditional_with_colon
# Logical short-circuit-like behavior influences status in composite arithmetic expressions.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if (( 0 && (2/0) )); then
  fail "this branch should not run"
fi
if (( 1 || (2/0) )); then
  :
else
  fail "this branch should run"
fi
(( (0 || 1) ? 7 : 0 )); st=$?
[ "$st" -eq 0 ] || fail "7 branch should succeed"
echo PASS
exit 0

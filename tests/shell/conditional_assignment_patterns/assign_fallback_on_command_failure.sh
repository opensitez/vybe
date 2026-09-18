#!/usr/bin/env bash
# vybe-test: bash/conditional_assignment_patterns/assign_fallback_on_command_failure
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
v=$(false) || v=fallback
[ "$v" = fallback ] || fail "got [$v]"
w=$(echo real) || w=fallback
[ "$w" = real ] || fail "got [$w]"
echo PASS
exit 0

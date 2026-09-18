#!/usr/bin/env bash
# vybe-test: bash/command_terminators_and_separators/and_or_same_precedence_left_associative
# "true || echo A && echo B" is "(true || echo A) && echo B", so B prints.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(true || echo A && echo B)
[ "$out" = B ] || fail "want [B] got [$out]"
out=$(false && echo C || echo D)
[ "$out" = D ] || fail "want [D] got [$out]"
echo PASS
exit 0

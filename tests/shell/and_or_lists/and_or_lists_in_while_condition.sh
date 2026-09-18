#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/and_or_lists_in_while_condition
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
i=0
while ((i < 1)) && :; do i=$((i + 1)); break; done
[ "$i" -eq 1 ] || fail "and in while condition should run body"
i=0
while false && ((i += 1)); do i=$((i + 1)); done
[ "$i" -eq 0 ] || fail "and-false while condition should never enter"
echo PASS
exit 0

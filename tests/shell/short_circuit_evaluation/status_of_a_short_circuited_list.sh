#!/usr/bin/env bash
# vybe-test: bash/short_circuit_evaluation/status_of_a_short_circuited_list
# When the right side is skipped, the list's status is the left side's.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
false && true; st=$?
[ "$st" -eq 1 ] || fail "false && true: want 1 got $st"
true || false; st=$?
[ "$st" -eq 0 ] || fail "true || false: want 0 got $st"
(exit 3) && true; st=$?
[ "$st" -eq 3 ] || fail "(exit 3) && true: want 3 got $st"
echo PASS
exit 0

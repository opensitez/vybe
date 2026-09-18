#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/zero_and_nonzero_results_map_to_failure_and_success
# Any non-zero arithmetic result is success; zero is failure.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(( 2 )); st=$?
[ "$st" -eq 0 ] || fail "(( 2 )) want 0 got $st"
(( -9 )); st=$?
[ "$st" -eq 0 ] || fail "(( -9 )) want 0 got $st"
(( 0 )); st=$?
[ "$st" -eq 1 ] || fail "(( 0 )) want 1 got $st"
[ $(( -1 )) -eq -1 ] || fail "-1 check"
echo PASS
exit 0

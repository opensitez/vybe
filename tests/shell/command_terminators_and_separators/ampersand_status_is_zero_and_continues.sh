#!/usr/bin/env bash
# vybe-test: bash/command_terminators_and_separators/ampersand_status_is_zero_and_continues
# Starting a background job succeeds (status 0) whatever the job later returns.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
false & st=$?
wait
[ "$st" = 0 ] || fail "launch status want 0 got $st"
echo PASS
exit 0

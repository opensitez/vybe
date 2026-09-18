#!/usr/bin/env bash
# vybe-test: bash/command_terminators_and_separators/semicolon_runs_next_regardless_of_status
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(false; echo after-false; true; echo after-true)
[ "$out" = $'after-false\nafter-true' ] || fail "got [$out]"
echo PASS
exit 0

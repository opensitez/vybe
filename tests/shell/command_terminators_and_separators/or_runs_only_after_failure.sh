#!/usr/bin/env bash
# vybe-test: bash/command_terminators_and_separators/or_runs_only_after_failure
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(true || echo skipped; false || echo taken)
[ "$out" = taken ] || fail "got [$out]"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/compound_command_syntax/subshell_in_pipeline_preserves_pipeline_status
# A subshell compound command inside a pipeline forwards its inner exit status to PIPESTATUS.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(exit 3) | cat >/dev/null
pipe_statuses=("${PIPESTATUS[@]}")
[ "${pipe_statuses[0]}" -eq 3 ] || fail "PIPESTATUS[0]: want 3, got ${pipe_statuses[0]}"
[ "${pipe_statuses[1]}" -eq 0 ] || fail "PIPESTATUS[1]: want 0, got ${pipe_statuses[1]}"
echo PASS
exit 0

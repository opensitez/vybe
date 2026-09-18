#!/usr/bin/env bash
# vybe-test: bash/command_terminators_and_separators/pipe_ampersand_redirects_stderr_and_stdout
# The |& separator connects both stdout and stderr of the command to the pipeline.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$( (echo "err_data" >&2) |& cat )
[ "$out" = "err_data" ] || fail "|& stderr capture: want 'err_data', got [$out]"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/exported_variables/export_variable_visible_to_child_process
# A variable exported with 'export' is placed in the environment and visible to child processes.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export PARENT_EXPORTED="payload_456"
child_seen=$( "$BASH" -c 'printf "%s\n" "$PARENT_EXPORTED"' )
[ "$child_seen" = "payload_456" ] || fail "child process failed to see exported variable: got [$child_seen]"
echo PASS
exit 0

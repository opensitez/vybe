#!/usr/bin/env bash
# vybe-test: bash/exported_variables/export_visibility_in_nested_child_processes
# Exported variables propagate transitively through multiple generations of spawned child processes.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export DEEP_CHAIN="transitive_export"
cmd='printf "%s\n" "$DEEP_CHAIN"'
deep_res=$(
    "$BASH" -c '"$BASH" -c "$1" _' _ "$cmd"
)
[ "$deep_res" = "transitive_export" ] || fail "transitive export propagation failed: got [$deep_res]"
echo PASS
exit 0

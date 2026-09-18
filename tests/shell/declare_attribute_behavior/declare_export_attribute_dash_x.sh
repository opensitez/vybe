#!/usr/bin/env bash
# vybe-test: bash/declare_attribute_behavior/declare_export_attribute_dash_x
# The 'declare -x' flag marks a variable for export into the environment of child processes.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -x EXPORTED_TOKEN="tok_9988"
res=$(
    # Child subshell environment inherits exported variable
    "$BASH" -c 'printf "%s\n" "$EXPORTED_TOKEN"'
)
[ "$res" = "tok_9988" ] || fail "declare -x failed to export to child process: got [$res]"
echo PASS
exit 0

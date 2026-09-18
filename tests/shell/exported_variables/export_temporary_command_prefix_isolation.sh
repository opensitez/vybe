#!/usr/bin/env bash
# vybe-test: bash/exported_variables/export_temporary_command_prefix_isolation
# Prefixing a command with 'VAR=val' exports VAR strictly to that command and not to later commands.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset TRANSIENT_FLAG
saw_first=$( TRANSIENT_FLAG="set_for_one" "$BASH" -c 'printf "%s\n" "$TRANSIENT_FLAG"' )
saw_second=$( "$BASH" -c 'printf "%s\n" "$TRANSIENT_FLAG"' )
[ "$saw_first" = "set_for_one" ] || fail "transient prefix failed for first command: got [$saw_first]"
[ -z "$saw_second" ] || fail "transient prefix leaked to second command: got [$saw_second]"
echo PASS
exit 0

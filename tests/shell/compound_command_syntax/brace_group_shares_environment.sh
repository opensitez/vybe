#!/usr/bin/env bash
# vybe-test: bash/compound_command_syntax/brace_group_shares_environment
# Unlike a subshell, a brace group { ...; } executes within the current shell environment.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x="parent"
{
    x="modified_in_brace"
}
[ "$x" = "modified_in_brace" ] || fail "brace group did not mutate environment: got [$x]"
echo PASS
exit 0

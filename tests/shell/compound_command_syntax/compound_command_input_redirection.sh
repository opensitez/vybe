#!/usr/bin/env bash
# vybe-test: bash/compound_command_syntax/compound_command_input_redirection.sh
# Redirection of standard input into a compound command provides input to all commands within it.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
{
    read -r first
    read -r second
} <<< $'alpha\nbeta'
[ "$first" = "alpha" ] || fail "first: want 'alpha', got [$first]"
[ "$second" = "beta" ] || fail "second: want 'beta', got [$second]"
echo PASS
exit 0

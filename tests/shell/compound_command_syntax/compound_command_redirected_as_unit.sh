#!/usr/bin/env bash
# vybe-test: bash/compound_command_syntax/compound_command_redirected_as_unit
# A redirection applied to a compound command applies to all commands executed inside its body.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
capture=$(
    {
        printf 'alpha\n'
        printf 'beta\n'
    } 2>&1
)
expected=$(printf 'alpha\nbeta')
[ "$capture" = "$expected" ] || fail "redirected compound command: got [$capture]"
echo PASS
exit 0

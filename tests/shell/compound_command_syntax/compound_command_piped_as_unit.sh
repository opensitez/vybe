#!/usr/bin/env bash
# vybe-test: bash/compound_command_syntax/compound_command_piped_as_unit
# A compound command can be connected directly to a pipeline as a single command unit.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$({
    printf 'line1\n'
    printf 'line2\n'
} | cat)
expected=$(printf 'line1\nline2')
[ "$out" = "$expected" ] || fail "piped brace group: got [$out]"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/command_terminators_and_separators/ampersand_semicolon_is_error
# & already terminates the command; a following ; has nothing to terminate.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval 'true &; echo a' 2>/dev/null; st=$?
[ "$st" -ne 0 ] || fail "&; must be a syntax error"
out=$(true & echo a; wait)
[ "$out" = a ] || fail "& followed directly by a command is fine, got [$out]"
echo PASS
exit 0

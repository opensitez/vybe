#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/evaluation_error_gives_status_1_and_shell_continues
# Unlike a failing $(( )) expansion, an error inside (( )) does not abort the
# shell: the command just fails.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(( 1/0 )) 2>/dev/null; st=$?
[ "$st" -eq 1 ] || fail "division by zero: want 1 got $st"
(( 1 + )) 2>/dev/null; st=$?
[ "$st" -eq 1 ] || fail "syntax error: want 1 got $st"
after=reached
[ "$after" = reached ] || fail "unreachable"
echo PASS
exit 0

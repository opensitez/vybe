#!/usr/bin/env bash
# vybe-test: bash/command_substitution/status_is_that_of_the_inner_command
# An assignment whose only expansion is a command substitution takes that
# substitution's exit status; exit inside it ends only the subshell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=$(exit 3); st=$?
[ "$st" -eq 3 ] || fail "want 3 got $st"
x=$(false; true); st=$?
[ "$st" -eq 0 ] || fail "last inner command wins, want 0 got $st"
x=$(exit 7)
[ "$?" -eq 7 ] || fail "exit in substitution must not exit the parent"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/process_substitution/dollar_bang_and_wait_track_the_substitution
# The command's own status ignores the substitution; $! names the
# substitution's process and wait retrieves its status.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: < <(exit 5); st=$?
[ "$st" -eq 0 ] || fail "command status must be :'s, got $st"
: < <(exit 5)
wait $!; ws=$?
[ "$ws" -eq 5 ] || fail "wait \$! want 5 got $ws"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/command_terminators_and_separators/list_status_is_last_command_status
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
false; true; a=$?
true; false; b=$?
(exit 7); c=$?
[ "$a" = 0 ] || fail "want 0 got $a"
[ "$b" = 1 ] || fail "want 1 got $b"
[ "$c" = 7 ] || fail "want 7 got $c"
echo PASS
exit 0

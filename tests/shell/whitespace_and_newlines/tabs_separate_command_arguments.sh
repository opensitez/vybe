#!/usr/bin/env bash
# vybe-test: bash/whitespace_and_newlines/tabs_separate_command_arguments
# Horizontal tabs act as standard whitespace separators between command arguments.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set --	alpha	beta		gamma
[ "$#" -eq 3 ] || fail "arg count with tabs: want 3, got $#"
[ "$1" = "alpha" ] || fail "arg 1: want 'alpha', got [$1]"
[ "$2" = "beta" ] || fail "arg 2: want 'beta', got [$2]"
[ "$3" = "gamma" ] || fail "arg 3: want 'gamma', got [$3]"
echo PASS
exit 0

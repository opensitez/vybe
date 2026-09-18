#!/usr/bin/env bash
# vybe-test: bash/background_list_execution/background_subshell_variable_isolation
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
( x=9; printf "$x" >/tmp/bash_bg_exec_isol_$PPID ) &
pid=$!
wait "$pid"
read -r got </tmp/bash_bg_exec_isol_$PPID
rm -f /tmp/bash_bg_exec_isol_$PPID
[ "$x" -eq 0 ] || fail "parent variable changed to $x"
[ "$got" = "9" ] || fail "child saw its own change"
echo PASS
exit 0

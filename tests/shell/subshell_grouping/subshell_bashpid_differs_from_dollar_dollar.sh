#!/usr/bin/env bash
# vybe-test: bash/subshell_grouping/subshell_bashpid_differs_from_dollar_dollar
# $$ remains the PID of the invoking main shell, whereas BASHPID expands to the subshell process ID.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
parent_pid=$$
read -r sub_dollar sub_bashpid < <( (printf '%s %s\n' "$$" "$BASHPID") )
[ "$sub_dollar" -eq "$parent_pid" ] || fail "\$\$ changed in subshell: want $parent_pid, got $sub_dollar"
[ "$sub_bashpid" -ne "$parent_pid" ] || fail "BASHPID did not change in subshell: parent was $parent_pid"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/special_parameters/special_param_dollar_dollar_preserved_in_subshell
# In Bash, $$ does not change in subshells; it retains the process ID of the parent shell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
parent_pid=$$
subshell_pid=$( ( echo "$$" ) )
[ "$subshell_pid" = "$parent_pid" ] || fail "\$\$ changed inside subshell: parent $parent_pid, subshell $subshell_pid"
echo PASS
exit 0

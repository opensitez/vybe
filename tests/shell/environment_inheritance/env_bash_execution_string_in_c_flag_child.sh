#!/usr/bin/env bash
# vybe-test: bash/environment_inheritance/env_bash_execution_string_in_c_flag_child
# The BASH_EXECUTION_STRING variable contains the exact command string passed via the -c option.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
exec_str=$( "$BASH" -c 'printf "%s\n" "$BASH_EXECUTION_STRING"' )
[ "$exec_str" = 'printf "%s\n" "$BASH_EXECUTION_STRING"' ] || fail "BASH_EXECUTION_STRING mismatch: got [$exec_str]"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/positional_parameters/pos_param_dollar_zero_contains_script_or_shell_name
# The special parameter $0 contains the name or invocation path of the currently executing script.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ -n "$0" ] || fail "\$0 is empty"
case "$0" in
    *pos_param_dollar_zero_contains_script_or_shell_name*|*bash*) : ;;
    *) fail "\$0 does not identify current script or shell: got [$0]" ;;
esac
echo PASS
exit 0

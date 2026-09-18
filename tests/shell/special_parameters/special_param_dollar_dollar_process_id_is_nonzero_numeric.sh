#!/usr/bin/env bash
# vybe-test: bash/special_parameters/special_param_dollar_dollar_process_id_is_nonzero_numeric
# The $$ parameter expands to the decimal process ID of the invoking shell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
pid=$$
[ "$pid" -gt 0 ] 2>/dev/null || fail "\$\$ is not a positive integer: got [$pid]"
echo PASS
exit 0

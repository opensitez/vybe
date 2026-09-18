#!/usr/bin/env bash
# vybe-test: bash/positional_parameters/pos_param_subshell_inherits_positional_parameters
# Subshells inherit the current positional parameters from the parent shell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "sub1" "sub2"
sub_res=$(
    (
        printf '%s,%s\n' "$1" "$2"
    )
)
[ "$sub_res" = "sub1,sub2" ] || fail "subshell failed to inherit positional parameters: got [$sub_res]"
echo PASS
exit 0

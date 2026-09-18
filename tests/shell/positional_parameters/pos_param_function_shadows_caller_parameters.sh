#!/usr/bin/env bash
# vybe-test: bash/positional_parameters/pos_param_function_shadows_caller_parameters
# Inside a function body, positional parameters $1, $2, etc., shadow the caller's parameters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "caller_p1" "caller_p2"
target_function() {
    [ "$#" -eq 1 ] || fail "function \$#: want 1, got $#"
    [ "$1" = "fn_arg" ] || fail "function \$1: want 'fn_arg', got [$1]"
}
target_function "fn_arg"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/positional_parameters/pos_param_caller_parameters_restored_after_function
# When a function returns, the caller's original positional parameters are completely restored.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "outer1" "outer2"
dummy_fn() {
    set -- "inner_only"
}
dummy_fn "called_with_arg"
[ "$#" -eq 2 ] || fail "caller parameter count not restored: want 2, got $#"
[ "$1" = "outer1" ] && [ "$2" = "outer2" ] || fail "caller parameters altered after function returned"
echo PASS
exit 0

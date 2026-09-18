#!/usr/bin/env bash
# vybe-test: bash/function_definition_syntax/function_shift_positional_arguments
# The 'shift' command inside a function shifts the function's own parameters, leaving caller's intact.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "top1" "top2"
shift_fn() {
    shift
    [ "$1" = "b" ] || fail "shifted func arg: want 'b', got [$1]"
}
shift_fn "a" "b"
[ "$1" = "top1" ] || fail "caller arg 1 modified by func shift: got [$1]"
echo PASS
exit 0

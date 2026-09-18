#!/usr/bin/env bash
# vybe-test: bash/function_definition_syntax/function_positional_parameters_and_arg_count
# Function invocations create their own positional parameters $1..$n and parameter count $#.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "script_arg1" "script_arg2"
test_args() {
    [ "$#" -eq 3 ] || fail "func arg count: want 3, got $#"
    [ "$1" = "f1" ] || fail "func arg 1: want 'f1', got [$1]"
    [ "$2" = "f2" ] || fail "func arg 2: want 'f2', got [$2]"
    [ "$3" = "f3" ] || fail "func arg 3: want 'f3', got [$3]"
}
test_args "f1" "f2" "f3"
[ "$#" -eq 2 ] || fail "parent arg count changed after function: got $#"
[ "$1" = "script_arg1" ] || fail "parent arg 1 changed after function: got [$1]"
echo PASS
exit 0

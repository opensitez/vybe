#!/usr/bin/env bash
# vybe-test: bash/function_definition_syntax/function_array_argument_passing
# Arrays passed via "${arr[@]}" expand to individual positional parameters inside the function.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fruits=( "honey crisp apple" "yellow banana" "sweet cherry" )
sum_args() {
    [ "$#" -eq 3 ] || fail "arg count: want 3, got $#"
    [ "$1" = "honey crisp apple" ] || fail "arg 1: got [$1]"
    [ "$2" = "yellow banana" ] || fail "arg 2: got [$2]"
    [ "$3" = "sweet cherry" ] || fail "arg 3: got [$3]"
}
sum_args "${fruits[@]}"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/positional_parameters/pos_param_slice_offset_and_length
# Slicing positional parameters via ${@:offset:length} extracts a sub-array of arguments.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "a" "b" "c" "d" "e"
slice_args() {
    [ "$#" -eq 2 ] || fail "slice count: want 2, got $#"
    [ "$1" = "b" ] && [ "$2" = "c" ] || fail "slice elements mismatch: got [$1], [$2]"
}
slice_args "${@:2:2}"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/function_definition_syntax/function_default_return_status_matches_last_command
# A function without an explicit 'return' or with 'return' without argument returns the status of last command.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
implicit_status_fn() {
    (exit 37)
}
implicit_status_fn
st1=$?
[ "$st1" -eq 37 ] || fail "implicit status: want 37, got $st1"

bare_return_fn() {
    (exit 88)
    return
}
bare_return_fn
st2=$?
[ "$st2" -eq 88 ] || fail "bare return status: want 88, got $st2"
echo PASS
exit 0

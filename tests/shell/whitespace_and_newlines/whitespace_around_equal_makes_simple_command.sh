#!/usr/bin/env bash
# vybe-test: bash/whitespace_and_newlines/whitespace_around_equal_makes_simple_command
# Spaces around '=' turn an intended assignment into a simple command invocation with arguments.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
dummy_fn() {
    [ "$#" -eq 2 ] || fail "arg count: want 2, got $#"
    [ "$1" = "=" ] || fail "arg 1: want '=', got [$1]"
    [ "$2" = "value" ] || fail "arg 2: want 'value', got [$2]"
}
dummy_fn = value
echo PASS
exit 0

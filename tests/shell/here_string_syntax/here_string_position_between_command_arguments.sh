#!/usr/bin/env bash
# vybe-test: bash/here_string_syntax/here_string_position_between_command_arguments
# A here-string can appear interspersed between standard command arguments without disrupting them.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
inspect_args() {
    [ "$#" -eq 2 ] || fail "arg count: want 2, got $#"
    [ "$1" = "arg1" ] || fail "arg 1: want 'arg1', got [$1]"
    [ "$2" = "arg2" ] || fail "arg 2: want 'arg2', got [$2]"
    read -r stdin_line
    [ "$stdin_line" = "middle_payload" ] || fail "stdin: want 'middle_payload', got [$stdin_line]"
}
inspect_args "arg1" <<< "middle_payload" "arg2"
echo PASS
exit 0

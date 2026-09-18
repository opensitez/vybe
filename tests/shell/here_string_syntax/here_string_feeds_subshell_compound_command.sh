#!/usr/bin/env bash
# vybe-test: bash/here_string_syntax/here_string_feeds_subshell_compound_command
# Attaching a here-string to a subshell ( ... ) <<< "data" supplies stdin to the entire subshell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
sub_result=$(
    (
        read -r val
        printf 'subshell_got:%s\n' "$val"
    ) <<< "payload_for_subshell"
)
[ "$sub_result" = "subshell_got:payload_for_subshell" ] || fail "here-string to subshell failed: got [$sub_result]"
echo PASS
exit 0

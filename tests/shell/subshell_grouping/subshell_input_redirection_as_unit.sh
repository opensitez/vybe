#!/usr/bin/env bash
# vybe-test: bash/subshell_grouping/subshell_input_redirection_as_unit
# Input redirection to a subshell feeds sequential input records to multiple commands in the subshell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(
    read -r a
    read -r b
    [ "$a" = "first_entry" ] || exit 1
    [ "$b" = "second_entry" ] || exit 2
) <<< $'first_entry\nsecond_entry'
st=$?
[ "$st" -eq 0 ] || fail "subshell input redirection failed: status $st"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/here_string_syntax/here_string_feeds_brace_group_compound_command
# Attaching a here-string to a brace group { ...; } <<< "data" supplies stdin while running in-place.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=""
{
    read -r x
} <<< "in_place_payload"
[ "$x" = "in_place_payload" ] || fail "here-string to brace group failed: got [$x]"
echo PASS
exit 0

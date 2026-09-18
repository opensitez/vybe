#!/usr/bin/env bash
# vybe-test: bash/here_string_syntax/here_string_position_at_start_of_command
# The <<< operator and its word can appear at the very beginning of a simple command.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
<<< "leading_stream" read -r captured
[ "$captured" = "leading_stream" ] || fail "leading here-string failed: got [$captured]"
echo PASS
exit 0

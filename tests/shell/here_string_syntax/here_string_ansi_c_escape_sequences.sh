#!/usr/bin/env bash
# vybe-test: bash/here_string_syntax/here_string_ansi_c_escape_sequences
# ANSI-C quoting $'...' in a here-string argument preserves control characters like tabs and newlines.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
{
    read -r line1
    read -r line2
} <<< $'first_tab\tval\nsecond_line'
[ "$line1" = $'first_tab\tval' ] || fail "line 1 tab failed: got [$line1]"
[ "$line2" = "second_line" ] || fail "line 2 failed: got [$line2]"
echo PASS
exit 0

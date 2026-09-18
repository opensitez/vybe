#!/usr/bin/env bash
# vybe-test: bash/newline_sensitive_constructs/newline_in_ifs_splits_multiline_output
# When IFS contains a newline, unquoted expansions split at newline boundaries.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
multiline="first line"$'\n'"second line"
IFS=$'\n'
set -- $multiline
[ "$#" -eq 2 ] || fail "arg count after IFS newline split: want 2, got $#"
[ "$1" = "first line" ] || fail "arg 1: want 'first line', got [$1]"
[ "$2" = "second line" ] || fail "arg 2: want 'second line', got [$2]"
echo PASS
exit 0

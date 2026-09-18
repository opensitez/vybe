#!/usr/bin/env bash
# vybe-test: bash/parameter_count_and_shift/shift_preserves_spaces_and_newlines_in_remaining
# Shifting arguments preserves embedded spaces and newlines in all remaining positional parameters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
nl_data="row1"$'\n'"row2"
set -- "discard_me" "embedded space" "$nl_data"
shift
[ "$#" -eq 2 ] || fail "count after shift: want 2, got $#"
[ "$1" = "embedded space" ] || fail "spaced argument corrupted: got [$1]"
[ "$2" = "$nl_data" ] || fail "newline argument corrupted: got [$2]"
echo PASS
exit 0

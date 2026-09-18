#!/usr/bin/env bash
# vybe-test: bash/brace_expansion/brace_numeric_sequence_ascending
# The {x..y} syntax expands to an inclusive ascending sequence of integer values.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- {1..5}
[ "$#" -eq 5 ] || fail "sequence count: want 5, got $#"
[ "$*" = "1 2 3 4 5" ] || fail "ascending sequence mismatch: got [$*]"
echo PASS
exit 0

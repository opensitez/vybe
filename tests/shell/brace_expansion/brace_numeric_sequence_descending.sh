#!/usr/bin/env bash
# vybe-test: bash/brace_expansion/brace_numeric_sequence_descending
# The {x..y} syntax where x > y expands to an inclusive descending sequence of integer values.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- {5..1}
[ "$#" -eq 5 ] || fail "descending count: want 5, got $#"
[ "$*" = "5 4 3 2 1" ] || fail "descending sequence mismatch: got [$*]"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/brace_expansion/brace_numeric_sequence_with_descending_step
# The {x..y..incr} syntax handles descending sequences with positive step values.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- {10..0..2}
[ "$#" -eq 6 ] || fail "descending step count: want 6, got $#"
[ "$*" = "10 8 6 4 2 0" ] || fail "descending stepped sequence mismatch: got [$*]"
echo PASS
exit 0

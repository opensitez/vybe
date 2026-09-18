#!/usr/bin/env bash
# vybe-test: bash/brace_expansion/brace_numeric_sequence_with_positive_step
# The {x..y..incr} syntax increments by the specified step value between terms.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- {0..10..2}
[ "$#" -eq 6 ] || fail "step count: want 6, got $#"
[ "$*" = "0 2 4 6 8 10" ] || fail "stepped sequence mismatch: got [$*]"
echo PASS
exit 0

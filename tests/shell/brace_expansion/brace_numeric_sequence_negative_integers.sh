#!/usr/bin/env bash
# vybe-test: bash/brace_expansion/brace_numeric_sequence_negative_integers
# The {x..y} syntax supports negative integer endpoints traversing zero.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- {-3..2}
[ "$#" -eq 6 ] || fail "negative sequence count: want 6, got $#"
[ "$*" = "-3 -2 -1 0 1 2" ] || fail "negative sequence mismatch: got [$*]"
echo PASS
exit 0

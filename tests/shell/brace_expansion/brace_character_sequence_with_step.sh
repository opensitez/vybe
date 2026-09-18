#!/usr/bin/env bash
# vybe-test: bash/brace_expansion/brace_character_sequence_with_step
# The {c1..c2..incr} syntax steps through the character sequence by the given increment.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- {a..m..3}
[ "$#" -eq 5 ] || fail "stepped character count: want 5, got $#"
[ "$*" = "a d g j m" ] || fail "stepped character sequence mismatch: got [$*]"
echo PASS
exit 0

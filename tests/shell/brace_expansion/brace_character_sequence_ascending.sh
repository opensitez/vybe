#!/usr/bin/env bash
# vybe-test: bash/brace_expansion/brace_character_sequence_ascending
# The {c1..c2} syntax expands to an ascending sequence of single characters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- {a..e}
[ "$#" -eq 5 ] || fail "character sequence count: want 5, got $#"
[ "$*" = "a b c d e" ] || fail "character sequence mismatch: got [$*]"
echo PASS
exit 0

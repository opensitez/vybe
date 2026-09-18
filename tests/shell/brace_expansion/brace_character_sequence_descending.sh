#!/usr/bin/env bash
# vybe-test: bash/brace_expansion/brace_character_sequence_descending
# The {c1..c2} syntax where c1 comes after c2 expands in descending alphabetical order.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- {e..a}
[ "$#" -eq 5 ] || fail "descending character count: want 5, got $#"
[ "$*" = "e d c b a" ] || fail "descending character sequence mismatch: got [$*]"
echo PASS
exit 0

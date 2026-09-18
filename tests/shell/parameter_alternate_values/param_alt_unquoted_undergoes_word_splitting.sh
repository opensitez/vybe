#!/usr/bin/env bash
# vybe-test: bash/parameter_alternate_values/param_alt_unquoted_undergoes_word_splitting
# An unquoted alternate value expansion undergoes word splitting on IFS characters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
flag="enabled"
count_words() {
    [ "$#" -eq 3 ] || fail "unquoted alternate word splitting failed: want 3 words, got $#"
    [ "$1" = "one" ] && [ "$2" = "two" ] && [ "$3" = "three" ] || fail "words mismatch"
}
count_words ${flag:+one two three}
echo PASS
exit 0

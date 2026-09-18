#!/usr/bin/env bash
# vybe-test: bash/parameter_defaults/param_default_unquoted_undergoes_word_splitting
# Unquoted ${var:-default} undergoes word splitting on IFS characters during word expansion.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset missing_list
count_words() {
    [ "$#" -eq 3 ] || fail "unquoted default word splitting failed: want 3 words, got $#"
    [ "$1" = "alpha" ] && [ "$2" = "beta" ] && [ "$3" = "gamma" ] || fail "words mismatch"
}
count_words ${missing_list:-alpha beta gamma}
echo PASS
exit 0

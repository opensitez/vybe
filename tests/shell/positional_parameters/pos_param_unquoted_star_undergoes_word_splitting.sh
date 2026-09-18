#!/usr/bin/env bash
# vybe-test: bash/positional_parameters/pos_param_unquoted_star_undergoes_word_splitting
# Unquoted $* undergoes word splitting on IFS characters when passed to commands.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "split me" "and me"
count_split() {
    [ "$#" -eq 4 ] || fail "unquoted \$* should split into 4 words: got $#"
}
count_split $*
echo PASS
exit 0

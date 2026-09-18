#!/usr/bin/env bash
# vybe-test: bash/variable_lookup_and_unset/var_lookup_unquoted_undergoes_word_splitting
# Referencing unquoted $var splits the variable value into multiple words based on IFS characters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
spaced="one two three"
count_args() { printf '%s\n' "$#"; }
c=$(count_args $spaced)
[ "$c" -eq 3 ] || fail "unquoted expansion argument count: want 3, got $c"
echo PASS
exit 0

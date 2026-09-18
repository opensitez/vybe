#!/usr/bin/env bash
# vybe-test: bash/positional_parameters/pos_param_star_expansion_single_word_with_ifs
# The "$*" parameter expansion expands to a single word separated by the first character of IFS.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "alpha" "beta" "gamma"
IFS=':'
joined="$*"
[ "$joined" = "alpha:beta:gamma" ] || fail "\$* with IFS=':': want 'alpha:beta:gamma', got [$joined]"
IFS=','
joined_comma="$*"
[ "$joined_comma" = "alpha,beta,gamma" ] || fail "\$* with IFS=',': want 'alpha,beta,gamma', got [$joined_comma]"
echo PASS
exit 0

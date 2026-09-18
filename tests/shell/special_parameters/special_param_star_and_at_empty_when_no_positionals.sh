#!/usr/bin/env bash
# vybe-test: bash/special_parameters/special_param_star_and_at_empty_when_no_positionals
# When there are no positional parameters, "$*" expands to an empty string and "$@" expands to 0 words.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set --
[ -z "$*" ] || fail "\$* should expand to empty when \$# is 0"
count_at() {
    [ "$#" -eq 0 ] || fail "\$@ should expand to 0 words: got $#"
}
count_at "$@"
echo PASS
exit 0

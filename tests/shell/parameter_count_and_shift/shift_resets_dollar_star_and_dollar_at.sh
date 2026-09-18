#!/usr/bin/env bash
# vybe-test: bash/parameter_count_and_shift/shift_resets_dollar_star_and_dollar_at
# Shifting immediately updates the expansions of "$*" and "$@" to match the remaining parameters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "first" "second" "third"
shift
star="$*"
[ "$star" = "second third" ] || fail "\$* did not update after shift: got [$star]"
count_remaining() {
    [ "$#" -eq 2 ] || fail "\$@ word count after shift: want 2, got $#"
}
count_remaining "$@"
echo PASS
exit 0

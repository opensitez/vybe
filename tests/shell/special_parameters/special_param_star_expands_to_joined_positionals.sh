#!/usr/bin/env bash
# vybe-test: bash/special_parameters/special_param_star_expands_to_joined_positionals
# The $* parameter expands to the positional parameters joined by the first char of IFS.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "one" "two" "three"
IFS='|'
joined="$*"
[ "$joined" = "one|two|three" ] || fail "\$* joining with IFS='|' failed: got [$joined]"
echo PASS
exit 0

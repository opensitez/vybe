#!/usr/bin/env bash
# vybe-test: bash/special_parameters/special_param_length_expansion_on_hash
# The ${##} expansion calculates the character length of the positional parameter count $#.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "one" "two" "three"
[ "${##}" -eq 1 ] || fail "length of count 3: want 1, got ${##}"

set -- 1 2 3 4 5 6 7 8 9 10 11 12
[ "${##}" -eq 2 ] || fail "length of count 12: want 2, got ${##}"
echo PASS
exit 0

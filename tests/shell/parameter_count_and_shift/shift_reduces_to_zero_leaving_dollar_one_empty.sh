#!/usr/bin/env bash
# vybe-test: bash/parameter_count_and_shift/shift_reduces_to_zero_leaving_dollar_one_empty
# Shifting all parameters away sets $# to 0 and causes $1 to expand to an empty string.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "only_one"
shift
[ "$#" -eq 0 ] || fail "count should be 0: got $#"
[ -z "$1" ] || fail "\$1 should be empty after shifting last parameter: got [$1]"
echo PASS
exit 0

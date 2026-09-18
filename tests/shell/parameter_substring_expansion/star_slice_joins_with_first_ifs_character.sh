#!/usr/bin/env bash
# vybe-test: bash/parameter_substring_expansion/star_slice_joins_with_first_ifs_character
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- a b c d
IFS=-
[ "${*:2:2}" = "b-c" ] || fail "want [b-c] got [${*:2:2}]"
[ "${*: -2}" = "c-d" ] || fail "want [c-d] got [${*: -2}]"
echo PASS
exit 0

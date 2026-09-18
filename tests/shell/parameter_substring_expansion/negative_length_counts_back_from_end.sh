#!/usr/bin/env bash
# vybe-test: bash/parameter_substring_expansion/negative_length_counts_back_from_end
# A negative length is an end offset: ${x:1:-1} drops one char at each end.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=abcdef
[ "${x:1:-1}" = bcde ] || fail "want [bcde] got [${x:1:-1}]"
[ "${x:0:-1}" = abcde ] || fail "want [abcde] got [${x:0:-1}]"
[ "${x: -4:-1}" = cde ] || fail "both negative: want [cde] got [${x: -4:-1}]"
echo PASS
exit 0

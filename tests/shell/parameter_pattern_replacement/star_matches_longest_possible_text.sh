#!/usr/bin/env bash
# vybe-test: bash/parameter_pattern_replacement/star_matches_longest_possible_text
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=a.b.c
[ "${x/*./X}" = Xc ] || fail "want [Xc] got [${x/*./X}]"
[ "${x/.*/X}" = aX ] || fail "want [aX] got [${x/.*/X}]"
echo PASS
exit 0

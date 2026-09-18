#!/usr/bin/env bash
# vybe-test: bash/parameter_prefix_suffix_removal/percent_shortest_and_double_percent_longest_suffix
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=a.b.c
[ "${x%.*}" = a.b ] || fail "shortest: want [a.b] got [${x%.*}]"
[ "${x%%.*}" = a ] || fail "longest: want [a] got [${x%%.*}]"
echo PASS
exit 0

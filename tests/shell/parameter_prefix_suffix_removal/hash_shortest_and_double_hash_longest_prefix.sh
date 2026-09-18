#!/usr/bin/env bash
# vybe-test: bash/parameter_prefix_suffix_removal/hash_shortest_and_double_hash_longest_prefix
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=a.b.c
[ "${x#*.}" = b.c ] || fail "shortest: want [b.c] got [${x#*.}]"
[ "${x##*.}" = c ] || fail "longest: want [c] got [${x##*.}]"
echo PASS
exit 0

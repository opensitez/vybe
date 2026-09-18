#!/usr/bin/env bash
# vybe-test: bash/parameter_prefix_suffix_removal/pattern_is_anchored_to_the_edge
# # only removes a match at the very start and % only at the very end; a
# match in the middle, or no match at all, leaves the value unchanged.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=abc
[ "${x#b}" = abc ] || fail "b is not a prefix: got [${x#b}]"
[ "${x%b}" = abc ] || fail "b is not a suffix: got [${x%b}]"
[ "${x#z}" = abc ] || fail "no match: got [${x#z}]"
[ "${x#a}" = bc ] || fail "real prefix: got [${x#a}]"
echo PASS
exit 0

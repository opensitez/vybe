#!/usr/bin/env bash
# vybe-test: bash/parameter_pattern_replacement/hash_and_percent_anchor_the_match
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=aXa
[ "${x/#a/1}" = 1Xa ] || fail "prefix anchor: got [${x/#a/1}]"
[ "${x/%a/2}" = aX2 ] || fail "suffix anchor: got [${x/%a/2}]"
[ "${x/#X/-}" = aXa ] || fail "anchored pattern must not match in the middle: got [${x/#X/-}]"
echo PASS
exit 0

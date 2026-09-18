#!/usr/bin/env bash
# vybe-test: bash/parameter_prefix_suffix_removal/extglob_patterns_in_removal
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s extglob
x=123abc456
[ "${x##+([0-9])}" = abc456 ] || fail "+([0-9]) prefix: got [${x##+([0-9])}]"
[ "${x%%+([0-9])}" = 123abc ] || fail "+([0-9]) suffix: got [${x%%+([0-9])}]"
[ "${x#@(12|xy)}" = 3abc456 ] || fail "@(alt) prefix: got [${x#@(12|xy)}]"
echo PASS
exit 0

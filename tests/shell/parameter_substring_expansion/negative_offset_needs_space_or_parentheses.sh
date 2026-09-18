#!/usr/bin/env bash
# vybe-test: bash/parameter_substring_expansion/negative_offset_needs_space_or_parentheses
# ${x:-2} is the default-value operator, so a negative offset must be
# written ${x: -2} or ${x:(-2)}.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=abcdef
[ "${x: -2}" = ef ] || fail "space form: got [${x: -2}]"
[ "${x:(-2)}" = ef ] || fail "paren form: got [${x:(-2)}]"
[ "${x:-2}" = abcdef ] || fail ":-2 is a default expansion and yields x itself, got [${x:-2}]"
[ "${x: -3:2}" = de ] || fail "negative offset with length: got [${x: -3:2}]"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/parameter_substring_expansion/offset_and_length_are_arithmetic_expressions
# Both fields are evaluated as arithmetic: variables need no $, operators and
# nested expansions such as ${#x} are allowed, and surrounding spaces are fine.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=abcdef; i=1; j=1
[ "${x:i+1:j*2}" = cd ] || fail "expressions: got [${x:i+1:j*2}]"
[ "${x:${#x}-2}" = ef ] || fail "length-based offset: got [${x:${#x}-2}]"
[ "${x: 1 : 2 }" = bc ] || fail "spaces: got [${x: 1 : 2 }]"
echo PASS
exit 0

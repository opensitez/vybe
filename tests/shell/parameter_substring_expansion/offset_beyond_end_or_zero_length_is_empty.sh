#!/usr/bin/env bash
# vybe-test: bash/parameter_substring_expansion/offset_beyond_end_or_zero_length_is_empty
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=abc
[ -z "${x:3}" ] || fail "offset == length: got [${x:3}]"
[ -z "${x:10}" ] || fail "offset past end: got [${x:10}]"
[ -z "${x:1:0}" ] || fail "zero length: got [${x:1:0}]"
[ -z "${x: -10}" ] || fail "negative offset before the start: got [${x: -10}]"
echo PASS
exit 0

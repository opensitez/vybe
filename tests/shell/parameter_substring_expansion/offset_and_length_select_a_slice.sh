#!/usr/bin/env bash
# vybe-test: bash/parameter_substring_expansion/offset_and_length_select_a_slice
# Length counts characters from the offset and is clipped at the end.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=abcdef
[ "${x:1:3}" = bcd ] || fail "want [bcd] got [${x:1:3}]"
[ "${x:4:10}" = ef ] || fail "length past the end is clipped, got [${x:4:10}]"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/parameter_substring_expansion/offset_only_returns_rest_of_string
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=abcdef
[ "${x:2}" = cdef ] || fail "want [cdef] got [${x:2}]"
[ "${x:0}" = abcdef ] || fail "offset 0 is the whole string, got [${x:0}]"
echo PASS
exit 0

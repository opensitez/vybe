#!/usr/bin/env bash
# vybe-test: bash/backslash_escaping/backslash_before_newline_inside_variable_value
# A backslash stored in a variable is data: expanding it does not escape anything.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count() { echo $#; }
v='a\ b'
n=$(count $v)
[ "$n" = 2 ] || fail "backslash from expansion must not quote the space, got $n args"
[ "${#v}" -eq 4 ] || fail "value keeps its backslash, length want 4 got ${#v}"
echo PASS
exit 0

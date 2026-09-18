#!/usr/bin/env bash
# vybe-test: bash/parameter_substring_expansion/empty_length_field_is_empty_result
# ${x:1:} has an empty length expression, which evaluates to 0.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=abcdef
[ -z "${x:1:}" ] || fail "got [${x:1:}]"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/parameter_substring_expansion/unset_variable_slices_to_empty
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset nothing
[ -z "${nothing:1:2}" ] || fail "got [${nothing:1:2}]"
[ -z "${nothing: -1}" ] || fail "got [${nothing: -1}]"
echo PASS
exit 0

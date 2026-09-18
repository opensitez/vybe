#!/usr/bin/env bash
# vybe-test: bash/variable_naming_rules/var_name_read_validates_identifier
# The read builtin validates destination variable names and exits non-zero on invalid identifier names.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
read -r 1bad_var <<< "data" 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "read into 1bad_var should exit non-zero"
echo PASS
exit 0

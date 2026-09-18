#!/usr/bin/env bash
# vybe-test: bash/dynamic_variable_names/dyn_declare_rejects_invalid_constructed_name
# Dynamically passing an invalid assignment string to declare fails with a non-zero exit status.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
bad_assignment="99_slot=value"
declare "$bad_assignment" 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "declare with invalid dynamic name should fail"
echo PASS
exit 0

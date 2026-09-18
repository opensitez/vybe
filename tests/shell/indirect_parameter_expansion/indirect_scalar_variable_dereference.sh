#!/usr/bin/env bash
# vybe-test: bash/indirect_parameter_expansion/indirect_scalar_variable_dereference
# The ${!ptr} expansion dereferences ptr to retrieve the value of the variable named by ptr.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
underlying="resolved_payload"
ptr="underlying"
[ "${!ptr}" = "resolved_payload" ] || fail "indirect dereference failed: got [${!ptr}]"
echo PASS
exit 0

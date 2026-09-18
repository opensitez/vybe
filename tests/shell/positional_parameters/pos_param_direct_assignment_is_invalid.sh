#!/usr/bin/env bash
# vybe-test: bash/positional_parameters/pos_param_direct_assignment_is_invalid
# Direct assignment to a positional parameter like '1=val' is a syntax error / invalid identifier.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval '1="illegal"' 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "direct assignment to positional parameter '1=val' should fail"
echo PASS
exit 0

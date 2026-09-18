#!/usr/bin/env bash
# vybe-test: bash/variable_naming_rules/var_name_containing_colon_is_invalid
# Variable identifiers containing colons are invalid in variable assignments.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval 'foo:bar=invalid' 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "foo:bar=invalid should fail"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/variable_naming_rules/var_name_containing_hyphen_is_invalid
# Variable identifiers containing hyphens are invalid in Bash.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval 'foo-bar=invalid' 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "foo-bar=invalid should fail as hyphen is invalid in identifier"
echo PASS
exit 0

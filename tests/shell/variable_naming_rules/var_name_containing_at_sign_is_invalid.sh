#!/usr/bin/env bash
# vybe-test: bash/variable_naming_rules/var_name_containing_at_sign_is_invalid
# Variable identifiers containing '@' are invalid (except the special parameter $@ itself).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval 'user@host=invalid' 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "user@host=invalid should fail"
echo PASS
exit 0

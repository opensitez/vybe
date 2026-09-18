#!/usr/bin/env bash
# vybe-test: bash/declare_attribute_behavior/declare_dash_p_nonexistent_variable_fails
# Running 'declare -p' on a variable that does not exist produces a non-zero exit status.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset nonexistent_attribute_probe
declare -p nonexistent_attribute_probe 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "declare -p on nonexistent variable should return non-zero exit code"
echo PASS
exit 0

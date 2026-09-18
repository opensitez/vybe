#!/usr/bin/env bash
# vybe-test: bash/variable_lookup_and_unset/var_unset_removes_variable
# The unset builtin deletes a variable such that [[ -v var ]] tests false and $var is empty.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
temp_entry="present"
[[ -v temp_entry ]] || fail "temp_entry should initially be set"
unset temp_entry
[[ ! -v temp_entry ]] || fail "temp_entry should be unset after unset builtin"
[ -z "$temp_entry" ] || fail "unset variable value should be empty: got [$temp_entry]"
echo PASS
exit 0

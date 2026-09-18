#!/usr/bin/env bash
# vybe-test: bash/function_definition_syntax/function_unset_dash_f_removes_definition
# The 'unset -f name' command deletes a function definition from the shell environment.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
temp_fn() { printf 'exists\n'; }
[ "$(temp_fn)" = "exists" ] || fail "temp_fn setup failed"
unset -f temp_fn
type temp_fn >/dev/null 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "unset -f failed to remove function definition"
echo PASS
exit 0

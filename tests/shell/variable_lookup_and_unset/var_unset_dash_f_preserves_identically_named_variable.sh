#!/usr/bin/env bash
# vybe-test: bash/variable_lookup_and_unset/var_unset_dash_f_preserves_identically_named_variable
# 'unset -f name' removes the function without deleting a variable that shares the same name.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
coexisting="surviving_variable"
coexisting() { printf 'doomed_function\n'; }
unset -f coexisting
[ "$coexisting" = "surviving_variable" ] || fail "variable was destroyed by unset -f: got [$coexisting]"
type coexisting >/dev/null 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "function was not removed by unset -f"
echo PASS
exit 0

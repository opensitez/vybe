#!/usr/bin/env bash
# vybe-test: bash/subshell_grouping/subshell_function_definition_isolation
# Defining a shell function inside a subshell does not introduce it into the parent shell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset -f child_func 2>/dev/null
(
    child_func() { printf 'inside\n'; }
    [ "$(child_func)" = "inside" ] || exit 1
)
type child_func >/dev/null 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "function child_func should not exist in parent shell"
echo PASS
exit 0

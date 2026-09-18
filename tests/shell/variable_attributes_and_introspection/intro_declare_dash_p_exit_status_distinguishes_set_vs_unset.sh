#!/usr/bin/env bash
# vybe-test: bash/variable_attributes_and_introspection/intro_declare_dash_p_exit_status_distinguishes_set_vs_unset
# The 'declare -p' builtin exits with status 0 for defined variables and status 1 for unset variables.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
defined_slot="defined"
declare -p defined_slot >/dev/null 2>&1
st_defined=$?
[ "$st_defined" -eq 0 ] || fail "declare -p on defined variable: want 0, got $st_defined"

unset missing_slot
declare -p missing_slot >/dev/null 2>&1
st_missing=$?
[ "$st_missing" -ne 0 ] || fail "declare -p on unset variable should exit non-zero"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/variable_lookup_and_unset/var_unset_special_parameter_fails
# Executing unset on special parameters like '$#' or '$$' does not modify their values.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "alpha" "beta"
unset '#' 2>/dev/null
[ "$#" -eq 2 ] || fail "parameter count \$# modified by unset: want 2, got $#"

orig_pid=$$
unset '$' 2>/dev/null
[ "$$" -eq "$orig_pid" ] || fail "shell pid \$\$ modified by unset: got $$"
echo PASS
exit 0

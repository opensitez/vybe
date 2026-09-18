#!/usr/bin/env bash
# vybe-test: bash/variable_naming_rules/var_name_printf_dash_v_validates_identifier
# The 'printf -v varname' builtin validates the destination variable name, rejecting invalid identifiers.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
printf -v valid_dest "success_data"
[ "$valid_dest" = "success_data" ] || fail "valid printf -v failed"

printf -v "1_bad_dest" "bad_data" 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "printf -v with invalid identifier should return non-zero status"
echo PASS
exit 0

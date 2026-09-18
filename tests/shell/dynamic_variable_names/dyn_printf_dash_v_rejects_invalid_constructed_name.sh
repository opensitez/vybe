#!/usr/bin/env bash
# vybe-test: bash/dynamic_variable_names/dyn_printf_dash_v_rejects_invalid_constructed_name
# Constructing an invalid variable identifier (e.g. starting with digit or hyphen) causes printf -v to fail.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
bad_name="123_invalid"
printf -v "$bad_name" "value" 2>/dev/null
st1=$?
[ "$st1" -ne 0 ] || fail "printf -v with name starting with digit should fail"

bad_hyphen="hyphen-name"
printf -v "$bad_hyphen" "value" 2>/dev/null
st2=$?
[ "$st2" -ne 0 ] || fail "printf -v with hyphenated name should fail"
echo PASS
exit 0

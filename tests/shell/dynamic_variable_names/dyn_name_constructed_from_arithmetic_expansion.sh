#!/usr/bin/env bash
# vybe-test: bash/dynamic_variable_names/dyn_name_constructed_from_arithmetic_expansion
# Dynamic variable identifiers can be constructed using inline arithmetic expansions $(( ... )).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
printf -v "offset_$(( 5 * 5 ))" "%s" "twenty_five"
ptr="offset_25"
[ "${!ptr}" = "twenty_five" ] || fail "arithmetic-constructed variable name failed: got [${!ptr}]"
echo PASS
exit 0

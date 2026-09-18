#!/usr/bin/env bash
# vybe-test: bash/variable_lookup_and_unset/var_lookup_dollar_syntax
# Standard $var syntax retrieves the scalar value of the variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
sample_data="active_payload"
[ "$sample_data" = "active_payload" ] || fail "standard lookup failed: got [$sample_data]"
echo PASS
exit 0

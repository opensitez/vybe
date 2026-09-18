#!/usr/bin/env bash
# vybe-test: bash/nameref_variables/nameref_reads_underlying_target_variable
# Reading a nameref variable returns the value of the referenced target variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
source_data="active_data"
declare -n ref=source_data
[ "$ref" = "active_data" ] || fail "nameref read failed: got [$ref]"
echo PASS
exit 0

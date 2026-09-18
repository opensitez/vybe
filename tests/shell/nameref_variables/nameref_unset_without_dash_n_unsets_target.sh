#!/usr/bin/env bash
# vybe-test: bash/nameref_variables/nameref_unset_without_dash_n_unsets_target
# In Bash, invoking 'unset ref' (without -n) unsets the underlying target variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
victim_data="doomed_value"
declare -n ref=victim_data
unset ref
[[ ! -v victim_data ]] || fail "underlying target variable should be unset by unset ref"
echo PASS
exit 0

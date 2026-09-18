#!/usr/bin/env bash
# vybe-test: bash/nameref_variables/nameref_unset_dash_n_removes_ref_preserves_target
# The 'unset -n ref' command unsets the nameref itself, leaving the target variable and its value intact.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
target_data="intact_value"
declare -n ref=target_data
unset -n ref
[[ ! -R ref ]] || fail "ref still has nameref attribute after unset -n"
[ "$target_data" = "intact_value" ] || fail "target data was destroyed by unset -n: got [$target_data]"
echo PASS
exit 0

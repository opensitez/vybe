#!/usr/bin/env bash
# vybe-test: bash/variable_attributes_and_introspection/intro_test_dash_v_evaluates_associative_array_key_existence
# The [[ -v map[key] ]] expression introspects whether a specific associative array key is present.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A settings=( [timeout]="30" [debug]="" )
[[ -v settings[timeout] ]] || fail "key 'timeout' should exist"
[[ -v settings[debug] ]] || fail "empty key 'debug' should exist"
[[ ! -v settings[retries] ]] || fail "unassigned key 'retries' should not exist"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/nameref_variables/nameref_function_pass_by_reference_associative_array
# A function using 'local -n' receives an associative array and sets keys in caller's map.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set_config_entry() {
    local -n map_ref=$1
    local k=$2
    local v=$3
    map_ref[$k]="$v"
}
declare -A app_config=( [env]="prod" )
set_config_entry app_config "version" "2.1.0"
[ "${app_config[version]}" = "2.1.0" ] || fail "pass-by-reference associative array insertion failed"
[ "${app_config[env]}" = "prod" ] || fail "pre-existing map entry altered"
echo PASS
exit 0

#!/usr/bin/env bash
# vybe-test: bash/dynamic_variable_names/dyn_prefix_matching_star_operator
# The ${!prefix*} expansion expands to a space-separated list of variable names matching the prefix.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
app_cfg_port=8080
app_cfg_host="localhost"
app_cfg_ssl="true"
matches="${!app_cfg_*}"
case "$matches" in
    *"app_cfg_port"*"app_cfg_host"*"app_cfg_ssl"*|*"app_cfg_host"*"app_cfg_port"*"app_cfg_ssl"*) : ;;
    *) fail "prefix star expansion missing matching variables: got [$matches]" ;;
esac
echo PASS
exit 0

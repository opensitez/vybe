#!/usr/bin/env bash
# vybe-test: bash/indirect_parameter_expansion/indirect_prefix_matching_star_operator
# The ${!prefix*} expansion expands to a string containing variable names starting with prefix.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
cfg_host="127.0.0.1"
cfg_port="5432"
cfg_user="postgres"
matched="${!cfg_*}"
case "$matched" in
    *"cfg_host"*"cfg_port"*"cfg_user"*|*"cfg_host"*"cfg_user"*"cfg_port"*) : ;;
    *) fail "prefix star expansion missing variables: got [$matched]" ;;
esac
echo PASS
exit 0
